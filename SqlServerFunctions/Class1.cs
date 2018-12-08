using Microsoft.SqlServer.Server;
using System;
using System.Collections;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

//namespace Nskd
//{
public class SqlServerFunctions
{
    [SqlFunction(IsDeterministic = true)]
    public static SqlBoolean IsMatch(SqlString str, SqlString pattern)
    {
        SqlBoolean r = false;
        // Проверяем только если есть что проверять
        if (!str.IsNull)
        {
            if (!pattern.IsNull)
            {
                // если шаблон задан, то проверяем
                try
                {
                    Regex re = new Regex(pattern.ToString());
                    r = (SqlBoolean)re.IsMatch(str.ToString());
                }
                catch (Exception) { }
            }
            else
            {
                // если шаблон не задан, то проверять не надо. 
                // Всё годится кроме null, но null уже исключён.
                r = true;
            }
        }
        return r;
    }

    // функция возвращающая таблицу
    [SqlFunction(FillRowMethodName = "FillRow")]
    public static IEnumerable InitMethod(SqlString str, SqlString pattern)
    {
        // всегда возвращаем одну строку из десяти полей
        ArrayList a = new ArrayList(1);
        a.Add(null);
        if (!str.IsNull && !pattern.IsNull)
        {
            try
            {
                Regex re = new Regex(pattern.ToString());
                Match m = re.Match(str.ToString());
                a[0] = m.Groups;
            }
            catch (Exception) { }
        }
        return a;
    }

    public static void FillRow(
        Object obj,
        out SqlString g0,
        out SqlString g1,
        out SqlString g2,
        out SqlString g3,
        out SqlString g4,
        out SqlString g5,
        out SqlString g6,
        out SqlString g7,
        out SqlString g8,
        out SqlString g9)
    {
        GroupCollection groups = null;
        Int32 count = 0;
        if (obj != null)
        {
            groups = obj as GroupCollection;
            if (groups != null)
            {
                count = groups.Count;
            }
        }
        g0 = new SqlString((count > 0) ? groups[0].Value : String.Empty);
        g1 = new SqlString((count > 1) ? groups[1].Value : String.Empty);
        g2 = new SqlString((count > 2) ? groups[2].Value : String.Empty);
        g3 = new SqlString((count > 3) ? groups[3].Value : String.Empty);
        g4 = new SqlString((count > 4) ? groups[4].Value : String.Empty);
        g5 = new SqlString((count > 5) ? groups[5].Value : String.Empty);
        g6 = new SqlString((count > 6) ? groups[6].Value : String.Empty);
        g7 = new SqlString((count > 7) ? groups[7].Value : String.Empty);
        g8 = new SqlString((count > 8) ? groups[8].Value : String.Empty);
        g9 = new SqlString((count > 9) ? groups[9].Value : String.Empty);
    }

    // разбор строки описывающей дозировку в ГРЛС
    // возвращающает таблицу 
    // количество строк соответствует количеству действующих средств (разделитель - '+')
    // каждая строка содержит 4 поля
    // по два для числителя и знаменателя концентрации (число, ед. изм.)
    [SqlFunction(FillRowMethodName = "ParseDozFillRow")]
    public static IEnumerable ParseDoz(SqlString doz)
    {
        ArrayList rows = new ArrayList();
        if (!doz.IsNull)
        {
            String[] ps = doz.ToString().Split('+');
            for (int pi = 0; pi < ps.Length; pi++)
            {
                String p = ps[pi];
                rows.Add(p);
            }
        }
        return rows;
    }
    public static void ParseDozFillRow(
        Object row,
        out SqlDouble c1,
        out SqlString m1,
        out SqlDouble c2,
        out SqlString m2,
        out SqlDouble cp)
    {
        c1 = 0;
        m1 = String.Empty;
        c2 = 0;
        m2 = String.Empty;
        cp = 0;

        Regex re = new Regex(
            @"^" // начало
            + @"\s*(\d+[\d\., ]*)?" // число с фиксированной точкой (g1)
            + @"\s*(тыс|млн)?\.?" // множитель (g2)
            + @"\s*(%|\w*)?\.?" // единица измерения (g3)
        );
        String[] ud = ((String)row).Split(new Char[] { '/', '|', '\\' });
        for (int i = 0; i < Math.Min(ud.Length, 2); i++)
        {
            String t = ud[i];
            Match m = re.Match(t);
            GroupCollection gs = m.Groups;
            // три группы
            Double g1 = ((i == 0) ? 0 : 1);
            if (gs.Count > 1)
            {
                Double.TryParse(gs[1].Value.Replace(" ", String.Empty).Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out g1);
            }
            String g2 = gs.Count > 2 ? gs[2].Value : String.Empty;
            switch (g2)
            {
                case "тыс":
                    g1 *= 1000;
                    break;
                case "млн":
                    g1 *= 1000000;
                    break;
                default:
                    break;
            }
            String g3 = gs.Count > 3 ? gs[3].Value : String.Empty;
            if (i == 0)
            {
                c1 = new SqlDouble(g1);
                m1 = new SqlString(g3);
            }
            else
            {
                c2 = new SqlDouble(g1);
                m2 = new SqlString(g3);
                if (m2.Value == "мл")
                {
                    switch (m1.Value)
                    {
                        case "г":
                            if (g1 == 0) g1 = 1;
                            cp = (c1.Value * 100 / g1);
                            break;
                        case "мг":
                            if (g1 == 0) g1 = 1;
                            cp = (c1.Value / 10 / g1);
                            break;
                        case "мкг":
                            if (g1 == 0) g1 = 1;
                            cp = (c1.Value / 10000 / g1);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
//}
/* пример функции возвращающей таблицу
using System;  
using System.Data.Sql;  
using Microsoft.SqlServer.Server;  
using System.Collections;  
using System.Data.SqlTypes;  
using System.Diagnostics;  
  
public class TabularEventLog  
{  
    [SqlFunction(FillRowMethodName = "FillRow")]  
    public static IEnumerable InitMethod(String logname)  
    {  
        return new EventLog(logname).Entries;    }  
  
    public static void FillRow(Object obj, out SqlDateTime timeWritten, out SqlChars message, out SqlChars category, out long instanceId)  
    {  
        EventLogEntry eventLogEntry = (EventLogEntry)obj;  
        timeWritten = new SqlDateTime(eventLogEntry.TimeWritten);  
        message = new SqlChars(eventLogEntry.Message);  
        category = new SqlChars(eventLogEntry.Category);  
        instanceId = eventLogEntry.InstanceId;  
    }  
}  

 * 
use master;  
-- Replace SQL_Server_logon with your SQL Server user credentials.  
GRANT EXTERNAL ACCESS ASSEMBLY TO [SQL_Server_logon];   
-- Modify the following line to specify a different database.  
ALTER DATABASE master SET TRUSTWORTHY ON;  
  
-- Modify the next line to use the appropriate database.  
CREATE ASSEMBLY tvfEventLog   
FROM 'D:\assemblies\tvfEventLog\tvfeventlog.dll'   
WITH PERMISSION_SET = EXTERNAL_ACCESS;  
GO  
CREATE FUNCTION ReadEventLog(@logname nvarchar(100))  
RETURNS TABLE   
(logTime datetime,Message nvarchar(4000),Category nvarchar(4000),InstanceId bigint)  
AS   
EXTERNAL NAME tvfEventLog.TabularEventLog.InitMethod;  
GO  
 * 

 */
