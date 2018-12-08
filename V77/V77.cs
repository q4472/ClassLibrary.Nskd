using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Nskd.V77
{
    public class V77Object : IDisposable
    {
        private const BindingFlags INVOKE_METHOD = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static;
        private const BindingFlags GET_PROPERTY = BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Static;
        private const BindingFlags SET_PROPERTY = BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.Static;

        public Object ComObject { get; private set; }
        private Type oType;

        public V77Object() { ComObject = null; oType = null; }
        public V77Object(Object o)
        {
            if (o != null)
            {
                switch (o.GetType().Name)
                {
                    case "__ComObject":
                        ComObject = o;
                        break;
                    case "V77Object":
                        ComObject = ((V77Object)o).ComObject; ;
                        break;
                    default:
                        break;
                }
                oType = ComObject.GetType();
            }
        }

        public Object GetProperty(String name)
        {
            return oType?.InvokeMember(name, GET_PROPERTY, null, ComObject, null);
        }

        public Object InvokeMethod(String name)
        {
            return oType?.InvokeMember(name, INVOKE_METHOD, null, ComObject, null);
        }

        public Object InvokeMethod(String name, Object value)
        {
            return oType?.InvokeMember(name, INVOKE_METHOD, null, ComObject, new Object[] { value });
        }

        public Object InvokeMethod(String name, Object[] args)
        {
            return oType?.InvokeMember(name, INVOKE_METHOD, null, ComObject, args);
        }

        public void Dispose()
        {
            if (ComObject != null)
            {
                Marshal.ReleaseComObject(ComObject);
                ComObject = null;
                oType = null;
            }
        }
    }
}

namespace Nskd.V77
{
    public class GlobalContext : V77Object
    {
        private static Object GetContext(String cnString)
        {
            var v77SApplication = Type.GetTypeFromProgID("V77S.Application");
            var rootComObject = Activator.CreateInstance(v77SApplication);
            var root = new V77Object(rootComObject);
            Int32 code = (Int32)root.InvokeMethod("RMTrade");
            Object[] pars = new Object[] { code, cnString, "NO_SPLASH_SHOW" };
            var isConnected = (Boolean)root.InvokeMethod("Initialize", pars);
            if (!isConnected) { root.Dispose(); root = null; }
            return root;
        }

        public GlobalContext(String cnString) : base(GetContext(cnString)) { }

        // Системные.Функции/процедуры.Преобразование типов
        /*
        public String Строка(Object Параметр)
        {
            return (String)InvokeMethod("String", Параметр);
        }
        */

        // Системные.Функции/процедуры.Специальные
        public Object ПолучитьПустоеЗначение(Object Тип = null)
        {
            Object temp;
            if (Тип == null)
                temp = InvokeMethod("ПолучитьПустоеЗначение");
            else
                temp = InvokeMethod("ПолучитьПустоеЗначение", Тип);
            return temp;
        }
        public Double ПустоеЗначение(Object Значение)
        {
            return (Double)InvokeMethod("ПустоеЗначение", Значение);
        }
        public V77Object СоздатьОбъект(String ИмяАгрегатногоТипа)
        {
            return new V77Object(InvokeMethod("CreateObject", ИмяАгрегатногоТипа));
        }

        private Метаданные метаданные;
        public Метаданные Метаданные
        {
            get
            {
                if (метаданные == null)
                {
                    метаданные = new Метаданные(GetProperty("Метаданные"));
                }
                return метаданные;
            }
        }

        private Перечисление перечисление;
        public Перечисление Перечисление
        {
            get
            {
                if (перечисление == null)
                {
                    перечисление = new Перечисление(GetProperty("Перечисление"));
                }
                return перечисление;
            }
        }

        private Запрос запрос;
        public Запрос Запрос
        {
            get
            {
                if (запрос == null)
                {
                    запрос = new Запрос(СоздатьОбъект("Запрос"));
                }
                return запрос;
            }
        }

        new public void Dispose()
        {
            if (метаданные != null) { метаданные.Dispose(); метаданные = null; }
            if (перечисление != null) { перечисление.Dispose(); перечисление = null; }
            if (запрос != null) { запрос.Dispose(); запрос = null; }
            base.Dispose();
        }
    }

    public class Метаданные : V77Object
    {
        public Метаданные(Object o) : base(o) { }

        /*
            количество подчиненных объектов
            объект метаданных по указанному номеру
            объект метаданных по указанному идентификатору
        */
        public Double Справочник() { return (Double)InvokeMethod("Справочник"); }
        public МетаданныеСправочник Справочник(Int32 Номер) { return new МетаданныеСправочник(InvokeMethod("Справочник", Номер)); }
        public МетаданныеСправочник Справочник(String Идентификатор) { return new МетаданныеСправочник(InvokeMethod("Справочник", Идентификатор)); }

        public Double Перечисление() { return (Double)InvokeMethod("Перечисление"); }
        public МетаданныеПеречисление Перечисление(Int32 Номер) { return new МетаданныеПеречисление(InvokeMethod("Перечисление", Номер)); }
        public МетаданныеПеречисление Перечисление(String Идентификатор) { return new МетаданныеПеречисление(InvokeMethod("Перечисление", Идентификатор)); }

        public String ПолныйИдентификатор() { return (String)InvokeMethod("ПолныйИдентификатор"); }
    }

    public class МетаданныеСправочник : V77Object
    {
        public МетаданныеСправочник(Object o) : base(o) { }

        public String Идентификатор { get { return (String)GetProperty("Идентификатор"); } }
        /*
		- Идентификатор	"Товары"
		- Синоним
		- Комментарий	"Справочник товаров"
		- Владелец
		- КоличествоУровней	"3"
		- ДлинаКода	"10"
		- ДлинаНаименования	"100"
		- СерииКодов	"ВесьСправочник"
		- ТипКода	"Числовой"
		- ОсновноеПредставление	"ВВидеНаименования"
		- КонтрольУникальности	"1"
		- АвтоНумерация	"2"
		- ГруппыВпереди	"1"
		- СпособРедактирования	"ОбоимиСпособами"
		- ЕдинаяФормаЭлемента	"0"
		- ОсновнаяФорма	"Справочник.Товары.ФормаСписка.ФормаСписка"
		- ОсновнаяФормаДляВыбора	"Справочник.Товары.ФормаСписка.ФормаСписка"
		- ОбластьРаспространения	"ВсеИнформационныеБазы"
		- АвтоРегистрация	"1"
		- ДополнительныеКодыИБ
        */

        /*
            //[ROW_ID] [int]
            //[ID] [char](9)
            //[PARENTID] [char](9)  // Родитель
            //[CODE] [char](10)
            //[DESCR] [char](100)   // Наименование
            public String Наименование;
            //[ISFOLDER] [tinyint]
            //[ISMARK] [bit]        // ПометкаУдаления()
            //[VERSTAMP] [int]

        */

        public Double Реквизит() { return (Double)InvokeMethod("Реквизит"); }
        public МетаданныеСправочникРеквизит Реквизит(Int32 Номер) { return new МетаданныеСправочникРеквизит(InvokeMethod("Реквизит", Номер)); }
        public МетаданныеСправочникРеквизит Реквизит(String Идентификатор) { return new МетаданныеСправочникРеквизит(InvokeMethod("Реквизит", Идентификатор)); }

        public String ПолныйИдентификатор() { return (String)InvokeMethod("ПолныйИдентификатор"); }
    }

    public class МетаданныеСправочникРеквизит : V77Object
    {
        public МетаданныеСправочникРеквизит(Object o) : base(o) { }

        public String Идентификатор() { return (String)InvokeMethod("Идентификатор"); }

        public String ПолныйИдентификатор() { return (String)InvokeMethod("ПолныйИдентификатор"); }
    }

    public class МетаданныеПеречисление : V77Object
    {
        public МетаданныеПеречисление(Object o) : base(o) { }

        public МетаданныеПеречислениеЗначение Значение(Int32 ПорядковыйНомер)
        {
            return new МетаданныеПеречислениеЗначение(InvokeMethod("Значение", ПорядковыйНомер));
        }
    }

    public class МетаданныеПеречислениеЗначение : V77Object
    {
        public МетаданныеПеречислениеЗначение(Object o) : base(o) { }

        public String Представление()
        {
            return (String)InvokeMethod("Представление");
        }
    }

    public class Документ : V77Object
    {
        public Документ(Object o) : base(o) { }

        public DateTime ДатаДок { get { return (DateTime)ПолучитьАтрибут("ДатаДок"); } }
        public String НомерДок { get { return (String)ПолучитьАтрибут("НомерДок"); } }

        public String Вид()
        {
            return (String)InvokeMethod("Вид");
        }
        public Double ВыбратьДокументы(DateTime? Дата1 = null, DateTime? Дата2 = null)
        {
            return (Double)InvokeMethod("ВыбратьДокументы", new Object[] { Дата1, Дата2 });
        }
        public Double ВыбратьПодчиненныеДокументы(DateTime? Дата1 = null, DateTime? Дата2 = null, Документ Докум = null)
        {
            return (Double)InvokeMethod("ВыбратьПодчиненныеДокументы", new Object[] { Дата1, Дата2, Докум.ComObject });
        }
        public Double ВыбратьПоНомеру(String Номер, DateTime Дата, String ИдентВида = null)
        {
            return (Double)InvokeMethod("ВыбратьПоНомеру", new Object[] { Дата, ИдентВида });
        }
        public void Записать()
        {
            InvokeMethod("Записать");
        }
        public Double КоличествоСтрок()
        {
            return (Double)InvokeMethod("КоличествоСтрок");
        }
        public Double НайтиПоНомеру(String Номер, DateTime Дата, String ИдентВида = null)
        {
            return (Double)InvokeMethod("НайтиПоНомеру", new Object[] { Номер, Дата, ИдентВида });
        }
        public void НоваяСтрока()
        {
            InvokeMethod("НоваяСтрока");
        }
        public void Новый()
        {
            InvokeMethod("Новый");
        }
        public Object ПолучитьАтрибут(String ИмяРеквизита)
        {
            return InvokeMethod("ПолучитьАтрибут", ИмяРеквизита);
        }
        public Double ПолучитьДокумент()
        {
            return (Double)InvokeMethod("ПолучитьДокумент");
        }
        public void ПолучитьСтрокуПоНомеру(Int32 Номер)
        {
            InvokeMethod("ПолучитьСтрокуПоНомеру", Номер);
        }
        public Документ ТекущийДокумент()
        {
            return new Документ(InvokeMethod("ТекущийДокумент"));
        }
        public void УстановитьАтрибут(String ИмяРеквизита, Object Значение)
        {
            if (Значение is V77Object) Значение = ((V77Object)Значение).ComObject;
            InvokeMethod("УстановитьАтрибут", new Object[] { ИмяРеквизита, Значение });
        }
    }

    public class Справочник : V77Object
    {
        public Справочник(Object o) : base(o) { }

        public String Код { get { return (String)ПолучитьАтрибут("Код"); } }
        public String Наименование { get { return (String)ПолучитьАтрибут("Наименование"); } set { УстановитьАтрибут("Наименование", value); } }
        public Справочник Родитель { get { return new Справочник(ПолучитьАтрибут("Родитель")); } set { УстановитьАтрибут("Родитель", value); } }
        public Справочник Владелец { get { return new Справочник(ПолучитьАтрибут("Владелец")); } set { УстановитьАтрибут("Владелец", value); } }

        public Double Выбран()
        {
            return (Double)InvokeMethod("Выбран");
        }
        public Double ВыбратьЭлементы(Int32 Режим = 1)
        {
            return (Double)InvokeMethod("ВыбратьЭлементы", Режим);
        }
        public void Записать()
        {
            InvokeMethod("Записать");
        }
        public void ИспользоватьВладельца(Object Владелец, Int32 ФлагИзменения = 1)
        {
            if (Владелец is V77Object) Владелец = ((V77Object)Владелец).ComObject;
            InvokeMethod("ИспользоватьВладельца", new Object[] { Владелец, ФлагИзменения });
        }
        public void ИспользоватьРодителя(Object Родитель, Int32 ФлагИзменения = 1)
        {
            if (Родитель is V77Object) Родитель = ((V77Object)Родитель).ComObject;
            InvokeMethod("ИспользоватьРодителя", new Object[] { Родитель, ФлагИзменения });
        }
        public Double НайтиПоКоду(String Код, Int32 ФлагПоиска = -1)
        {
            if (ФлагПоиска == -1)
                return (Double)InvokeMethod("НайтиПоКоду", Код);
            else
                return (Double)InvokeMethod("НайтиПоКоду", new Object[] { Код, ФлагПоиска });
        }
        public Double НайтиПоНаименованию(String Наименование, Int32 Режим = 1, Int32 ФлагПоиска = 0)
        {
            return (Double)InvokeMethod("НайтиПоНаименованию", new Object[] { Наименование, Режим, ФлагПоиска });
        }
        public Double НайтиПоРеквизиту(String ИмяРеквизита, Object Значение, Int32 ФлагГлобальногоПоиска)
        {
            if (Значение is V77Object) Значение = ((V77Object)Значение).ComObject;
            return (Double)InvokeMethod("НайтиПоРеквизиту", new Object[] { ИмяРеквизита, Значение, ФлагГлобальногоПоиска });
        }
        public void НоваяГруппа()
        {
            InvokeMethod("НоваяГруппа");
        }
        public void Новый()
        {
            InvokeMethod("Новый");
        }
        public Object ПолучитьАтрибут(String ИмяРеквизита)
        {
            return InvokeMethod("ПолучитьАтрибут", ИмяРеквизита);
        }
        public Double ПолучитьЭлемент(Int32 Режим = 1)
        {
            return (Double)InvokeMethod("ПолучитьЭлемент", Режим);
        }
        public Double ПометкаУдаления()
        {
            return (Double)InvokeMethod("ПометкаУдаления");
        }
        public Справочник ТекущийЭлемент()
        {
            return new Справочник(InvokeMethod("ТекущийЭлемент"));
        }
        public void Удалить(Int32 Режим = 1)
        {
            InvokeMethod("Удалить", Режим);
        }
        public Double Уровень()
        {
            return (Double)InvokeMethod("Уровень");
        }
        public void УстановитьАтрибут(String ИмяРеквизита, Object Значение)
        {
            if (Значение is V77Object) Значение = ((V77Object)Значение).ComObject;
            InvokeMethod("УстановитьАтрибут", new Object[] { ИмяРеквизита, Значение });
        }
        public Double ЭтоГруппа()
        {
            return (Double)InvokeMethod("ЭтоГруппа");
        }

        // методы периодических реквизитов
        public Object Получить(String Дата = null)
        {
            return InvokeMethod("Получить", Дата);
        }
    }

    public class Перечисление : V77Object
    {
        public Перечисление(Object o) : base(o) { }

        // Методы вида перечисления
        public Перечисление ЗначениеПоИдентификатору(String Идентификатор)
        {
            return new Перечисление(InvokeMethod("ЗначениеПоИдентификатору", Идентификатор));
        }

        // Методы значения перечисления
        public Double Выбран()
        {
            return (Double)InvokeMethod("Выбран");
        }
        public String Идентификатор()
        {
            return (String)InvokeMethod("Идентификатор");
        }
        public String ПредставлениеВида()
        {
            return (String)InvokeMethod("ПредставлениеВида");
        }
        public Double ПорядковыйНомер()
        {
            return (Double)InvokeMethod("ПорядковыйНомер");
        }

        // Методы глобального атрибута "Перечисление"
        public Перечисление ПолучитьАтрибут(String ИмяВидаПеречисл)
        {
            return new Перечисление(InvokeMethod("ПолучитьАтрибут", ИмяВидаПеречисл));
        }
    }

    public class Периодический : V77Object
    {
        public Периодический(Object o) : base(o) { }

        public Object ДатаЗнач { get { return GetProperty("ДатаЗнач"); } }
        public Object Значение { get { return GetProperty("Значение"); } }
    }

    public class Запрос : V77Object
    {
        public Запрос(Object o) : base(o) { }

        public Double Выполнить(String ТекстЗапроса)
        {
            return (Double)InvokeMethod("Выполнить", ТекстЗапроса);
        }
        public Double Получить(Int32 ЗначениеГруппировки1) // <ЗначениеГруппировки1>,...,<ЗначениеГруппировкиN>
        {
            return (Double)InvokeMethod("Получить", ЗначениеГруппировки1);
        }
        public Object ПолучитьАтрибут(String ИмяАтрибута)
        {
            return InvokeMethod("ПолучитьАтрибут", ИмяАтрибута);
        }
        public Double Группировка()
        {
            return (Double)InvokeMethod("Группировка");
        }
        public Double Группировка(Int32 Группировка, Int32 Направление = 1)
        {
            return (Double)InvokeMethod("Группировка", new Object[] { Группировка, Направление });
        }
    }

    public class OcConvert
    {
        private static CultureInfo ic = CultureInfo.InvariantCulture;
        public static String ToD(Object value)
        {
            String result = "";
            switch (value.GetType().ToString())
            {
                case "System.String":
                    DateTime dt = new DateTime();
                    if (DateTime.TryParse((String)value, out dt))
                    {
                        result = dt.ToString("dd.MM.yyyy");
                    }
                    break;
                default:
                    break;
            }
            return result;
        }
        public static String ToN(Object value)
        {
            String result = "";
            switch (value.GetType().ToString())
            {
                case "System.String":
                    result = (new Regex(@"[^\d\.\,]")).Replace((String)value, "");
                    result = (new Regex(@",")).Replace(result, ".");
                    Double d = 0;
                    if (Double.TryParse(result, NumberStyles.Float, ic, out d))
                    {
                        result = d.ToString("0.####", ic);
                    }
                    break;
                default:
                    break;
            }
            return result;
        }
        public static String ToI(Object value)
        {
            String result = "";
            switch (value.GetType().ToString())
            {
                case "System.String":
                    result = (new Regex(@"[^\d]")).Replace((String)value, "");
                    break;
                default:
                    break;
            }
            return result;
        }
    }
}
