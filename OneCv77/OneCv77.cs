using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;

namespace Nskd.Oc77
{
    public class OcConnection
    {
        private static V77Connection[] v77cns = null;

        public static V77Connection Open(String ip = "192.168.135.77")
        {
            // Сюда попадают запросы от разных клиентов из разных потоков и процессов. 
            // Нет отслежевания сессии. 
            // При каждом запросе всё заново.
            // Однако процесс 1с запускаем только один раз. (пока не организован пул, где их буден несколько).

            // При первом запросе получаем новое подключение из пула.
            // При этом запускается 1с и таймер на 15 мин.
            // Если за этот период нет запросов от клиентов, то отключаемся от 1с.
            // При последующих запросах выдаём то подключение, которое уже есть.
            // Каждый запрос начинает отсчёт времени сначала.
            if ((v77cns == null) || (v77cns.Length == 0))
            {
                v77cns = new V77Connection[1];
                // Получить поключение из пула.
                v77cns[0] = V77ConnectionPool.GetV77Connection(ip);
            }
            else
            {
                // Подключение уже есть. Надо только перезапустить таймер.
                v77cns[0].ResetTimer();
            }
            return v77cns[0];
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

    public class V77Object
    {
        public Object ComObject { get; set; }
        public V77Object() { }
        public V77Object(Object comObject) { ComObject = comObject; }
    }

    public class V77System : V77Object
    {
        private static BindingFlags INVOKE_METHOD = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static;
        private Type rootComObjectType;

        public V77System(Object rootComObject)
            : base()
        {
            ComObject = rootComObject;
            rootComObjectType = rootComObject.GetType();
        }

        public V77Object CreateObject(String name)
        {
            V77Object obj = null;
            if (ComObject != null)
            {
                if (name.StartsWith("Документ.")) { obj = new V77Document(this); }
                else if (name.StartsWith("Справочник.")) { obj = new V77Reference(this); }
                else { obj = new V77Object(); }

                obj.ComObject = InvokeMethod(ComObject, "CreateObject", new Object[] { name });
            }
            return obj;
        }

        public Object InvokeMethod(Object comObject, String name)
        {
            return rootComObjectType.InvokeMember(name, INVOKE_METHOD, null, comObject, null);
        }

        public Object InvokeMethod(Object comObject, String name, Object value)
        {
            return rootComObjectType.InvokeMember(name, INVOKE_METHOD, null, comObject, new Object[] { value });
        }

        public Object InvokeMethod(Object comObject, String name, Object[] args)
        {
            return rootComObjectType.InvokeMember(name, INVOKE_METHOD, null, comObject, args);
        }
    }

    public class V77Document : V77Object
    {
        private V77System root;
        public V77Document(V77System _root)
            : base()
        {
            root = _root;
        }

        /// <summary>
        /// Синтаксис: НайтиПоНомеру("Номер", "Дата", "ИдентВида")
        /// Метод НайтиПоНомеру() позиционирует документ по номеру. 
        /// В качестве второго параметра задается любая дата из диапазона, в котором нужно искать документ с данным номером. 
        /// Поиск зависит от выбранного в конфигураторе способа уникальности номеров (по месяцу, году и др.).
        /// Метод может быть использован для объекта Документ общего вида, тогда для поиска нужно указать в параметре "ИдентВида" идентификатор вида документа или идентификатор Нумератора.
        /// Данный метод может использоваться только для объектов, созданных функцией СоздатьОбъект().
        /// </summary>
        /// <param name="num">Строковое выражение, содержащее значение номера искомого документа.</param>
        /// <param name="date">Выражение типа «дата».</param>
        /// <param name="kind">Необязательный параметр. Строковое выражение, содержащее идентификатор вида документа или идентификатор Нумератора.</param>
        /// <returns>
        /// Число 1 — если действие выполнено (документ найден);
        /// Число 0 — если действие не выполнено.
        /// </returns>
        public Double FindByNum(String num, DateTime date, String kind = "")
        {
            return (Double)root.InvokeMethod(ComObject, "НайтиПоНомеру", new Object[] { num, date, kind });
        }

        /// <summary>
        /// ТекущийДокумент()
        /// Метод ТекущийДокумент() возвращает значение позиционированного текущего документа (в целом, как объекта). 
        /// Данный метод применяется, например, если нужно документ передать как параметр в вызове какого-либо метода или присвоить какому-либо реквизиту.
        /// </summary>
        /// <returns>Значение текущего документа.</returns>
        public V77Object CurrentDocument()
        {
            return new V77Object(root.InvokeMethod(ComObject, "ТекущийДокумент"));
        }
    }

    public class V77Reference : V77Object
    {
        private V77System root;
        public V77Reference(V77System _root)
            : base()
        {
            root = _root;
        }

        /// <summary>
        /// Метод НайтиПоНаименованию() выполняет поиск элемента справочника по наименованию, 
        /// заданному параметром "descr" и позиционирует объект справочник на этом элементе.
        /// Данный метод может использоваться только для объектов, созданных функцией CreateObject().
        /// </summary>
        /// <param name="descr">
        /// Строковое выражение с наименованием искомого элемента справочника.
        /// </param>
        /// <param name="aria">
        /// Необязательный параметр. Числовое выражение — режим поиска:
        /// 1 — поиск внутри установленного подчинения (родителя);
        /// 0 — поиск во всем спра­вочнике вне зависимости от родителя.
        /// Значение по умолчанию — 1.
        /// </param>
        /// <param name="exactly">
        /// Необязательный параметр. Числовое выражение — флаг поиска:
        /// 1 — найти точное соответствие наиме­нования;
        /// 0 — найти наименование по первым сим­волам.
        /// Значение по умолчанию — 0.
        /// </param>
        /// <returns>
        /// Число 1 — если действие выполнено;
        /// Число 0 — если действие не выполнено (элемент не найден).
        /// </returns>
        public Double FindByDescr(String descr, Int32 aria = 1, Int32 exactly = 0)
        {
            return (Double)root.InvokeMethod(ComObject, "НайтиПоНаименованию", new Object[] { descr, aria, exactly });
        }

        public Double FindByCode(Object code, Int32 flag = 0)
        {
            return (Double)root.InvokeMethod(ComObject, "НайтиПоКоду", new Object[] { code }); // , flag
        }

        /// <summary>
        /// НайтиПоРеквизиту(ИмяРеквизита, Значение, ФлагГлобальногоПоиска)
        /// Метод НайтиПоРеквизиту() выполняет поиск первого элемента с указанным значением заданного реквизита и позиционирует объект справочник на этом элементе.
        /// Данный метод может использоваться только в том случае, если в конфигураторе при описании данного реквизита установлен признак «Сортировка» (Свойства реквизита — Дополнительные — Сортировка).
        /// Данный метод может использоваться только для объектов, созданных функцией СоздатьОбъект().
        /// </summary>
        /// <param name="name">Строковое выражение с наименованием реквизита.</param>
        /// <param name="comObject">Значение реквизита для поиска.</param>
        /// <param name="flag">Числовое выражение. Если 0, то поиск должен выполняться в пределах подчинения справочника, если 1, то поиск должен выполняться по всему справочнику.</param>
        /// <returns>
        /// Число 1 — если действие выполнено;
        /// Число 0 — если действие не выполнено (элемент не найден).
        /// </returns>
        public Double FindByAttribute(String name, Object comObject, Int32 flag)
        {
            return (Double)root.InvokeMethod(ComObject, "НайтиПоРеквизиту", new Object[] { name, comObject, flag });
        }

        public Double Selected()
        {
            return (Double)root.InvokeMethod(ComObject, "Выбран");
        }

        public V77Object CurrentItem()
        {
            return new V77Object(root.InvokeMethod(ComObject, "ТекущийЭлемент"));
        }

        /// <summary>
        /// Метод UseOwner() может применяться к объектам типа «справочник» в двух случаях:
        /// Для объектов, созданных функцией CreateObject(), метод UseOwner() устанавливает элемент справочника-владельца (которому подчинен текущий подчиненный справочник) в качестве параметра выборки. Данный метод используется до вызова метода ВыбратьЭлементы(), который фактически открывает выборку. Дальнейшая выборка при помощи метода ПолучитьЭлемент() будет происходить только среди тех элементов текущего подчиненного справочника, для которых владельцем является заданное значение элемента справочника-владельца <Владелец>. При записи нового элемента текущего справочника данный метод также задает владельца для нового элемента.
        /// Для объектов типа «справочник», которые являются реквизитами формы (например, в форме документа — реквизит документа типа «справочник») или реквизитами диалога (например, в форме отчета — реквизит диалога типа «справочник») метод ИспользоватьВладельца() позволяет программно установить некоторое значение справочника-владельца в качестве владельца, который будет использован при интерактивном выборе значения данного реквизита.
        /// </summary>
        /// <param name="owner">Необязательный параметр. Выражение со значением элемента справочника - владельца.</param>
        /// <param name="interactive">
        /// "ФлагИзменения"	Необязательный параметр. Этим флагом регулируется возможность интерактивного изменения владельца.
        /// 1 — пользователь может изменить владельца интерактивно,
        /// 0 — пользователь не может интерактивно изменить владельца.
        /// Этот параметр используется в случае использования данного метода для объектов типа «справочник», которые являются реквизитами формы или реквизитами диалога.
        /// </param>
        /// <returns>Значение элемента справочника-владельца для текущего подчиненного справочника (на момент до исполнения метода).</returns>
        public V77Object UseOwner(V77Object owner, Int32 interactive = 0)
        {
            return new V77Object(root.InvokeMethod(ComObject, "ИспользоватьВладельца", new Object[] { owner.ComObject, interactive }));
        }

        public V77Object UseParent(V77Object parent)
        {
            return new V77Object(root.InvokeMethod(ComObject, "ИспользоватьРодителя", parent.ComObject));
        }

        public void NewGroup()
        {
            root.InvokeMethod(ComObject, "НоваяГруппа");
        }

        public void New()
        {
            root.InvokeMethod(ComObject, "Новый");
        }

        public void Write()
        {
            root.InvokeMethod(ComObject, "Записать");
        }

        /// <summary>
        /// Удалить(Режим)
        /// Метод Удалить() удаляет (или делает пометку на удаление) текущий элемент или группу справочника.
        /// Данный метод может использоваться только для объектов, созданных функцией СоздатьОбъект().
        /// Замечание:
        /// Непосредственное удаление объекта следует применять очень аккуратно, так как это действие может нарушить ссылочную целостность информации. 
        /// Данный режим не рекомендуется использовать, если на данный объект могут быть ссылки в других объектах, например в реквизитах существующих документов.
        /// </summary>
        /// <param name="action">
        /// Режим	Числовое выражение:
        /// 1 — непосредственное удаление;
        /// 0 — пометка на удаление. Необязательный параметр.
        /// Значение по умолчанию — 1.
        /// </param>
        public void Delete(Int32 action = 1)
        {
            root.InvokeMethod(ComObject, "Удалить", action);
        }
        public Object GetAttrib(String name)
        {
            return root.InvokeMethod(ComObject, "ПолучитьАтрибут", name);
        }
        public void SetAttrib(String name, Object value)
        {
            root.InvokeMethod(ComObject, "УстановитьАтрибут", new Object[] { name, value });
        }
    }

    public class V77Connection
    {
        private static Type v77SApplication = Type.GetTypeFromProgID("V77S.Application");
        private static Object thisLock = new Object();
        private static Int32 timeout = 15 * 60 * 1000;

        private String cnString;
        private Timer timer;

        public Int32 ProcessId { get; set; }
        public V77System Root { get; set; }
        public Boolean IsConnected { get; set; }


        public V77Connection(String connectionString)
        {
            ProcessId = 0;
            List<Int32> pids = new List<Int32>();
            foreach (Process p in Process.GetProcesses()) { pids.Add(p.Id); }

            // получаем экземпляр COM объекта 1с.
            Object rootComObject = Activator.CreateInstance(v77SApplication);
            Root = new V77System(rootComObject);

            foreach (Process p in Process.GetProcesses())
            {
                if (!pids.Contains(p.Id) && (p.ProcessName.ToUpper() == "1CV7S"))
                {
                    ProcessId = p.Id;
                    break;
                }
            }

            IsConnected = false;
            cnString = connectionString;
            timer = new Timer(CloseByTimeout, null, Timeout.Infinite, Timeout.Infinite);
        }

        public void ResetTimer()
        {
            timer.Change(timeout, Timeout.Infinite);
        }

        private void CloseByTimeout(Object stateInfo)
        {
            Int32 hour = DateTime.Now.Hour;
            if ((hour < 7) || (hour >= 19))
            {
                // Останавливаем таймер и отключаесмя от 1с.
                timer.Change(Timeout.Infinite, Timeout.Infinite);
                Close();
            }
            else
            {
                // В рабочее время перезапускаем таймер.
                ResetTimer();
            }
        }

        public void Open()
        {
            if (Root.ComObject != null)
            {
                if (!IsConnected)
                {
                    lock (thisLock)
                    {
                        // код для режима 'предприятие'
                        Int32 code = (Int32)Root.InvokeMethod(Root.ComObject, "RMTrade");
                        Object[] pars = new Object[] { code, cnString, "NO_SPLASH_SHOW" };
                        // запускаем в режиме 'предприятие'
                        IsConnected = (Boolean)Root.InvokeMethod(Root.ComObject, "Initialize", pars);
                    }
                }
                ResetTimer();
            }
        }

        public void Close()
        {
            if ((Root.ComObject != null) && (IsConnected))
            {
                //Log.Write("11004", "Отключение от 1с.");
                Root.InvokeMethod(Root.ComObject, "ЗавершитьРаботуСистемы", 0);
            }
            IsConnected = false;
            // 12 секунд для завершения процесса отключения
            Thread.Sleep(12 * 1000);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Root.ComObject);
            //if (ProcessId != 0) { Process.GetProcessById(ProcessId).Kill(); }
            Root.ComObject = null;
        }
    }

    public class V77ConnectionPool
    {
        private static String folderG = @"\\SRV-TS2\dbase_1c$\Фармацея Фарм-Сиб";
        private static String folderL = @"C:\1c\FarmSib1C\";
        private static String userName = @"Соколов_Евгений_клиент_0"; // имя для пула подключений. номер будет потом заменён.
        private static String userPassword = @"yNFxfrvqxP";
        private static Int32 connectionsCount = 1;
        private static Boolean glFlag = true;
        private static String cnString;
        private static V77Connection[] pool;
        private static Boolean isReady;

        private static void fillPool(String ip)
        {
            // заполняем пул подключениями
            //for (int poolIndex = 0; poolIndex < pool.Length; poolIndex++)
            {
                // Номер клиента в имени пользователя соответствует номеру подключения.
                //cnString = (new Regex(@"клиент_\d*")).Replace(cnString, "клиент_" + (poolIndex + 1).ToString());

                // Номер клиента в имени пользователя временно соответствует ip. Для 77 - 1, а для 14 - 2
                switch (ip)
                {
                    case "192.168.135.14":
                        cnString = (new Regex(@"клиент_\d*")).Replace(cnString, "клиент_2");
                        break;
                    case "192.168.135.77":
                    default:
                        cnString = (new Regex(@"клиент_\d*")).Replace(cnString, "клиент_1"); //  + (poolIndex + 1).ToString()
                        break;
                }
                // Запускаем процесс 1с, но подключения пока нет.
                pool[0] = new V77Connection(cnString);
            }
            isReady = true;
        }

        static V77ConnectionPool()
        {
            String folder = glFlag ? folderG : folderL;
            cnString = "/d\"" + folder + "\" /n" + userName + " /p" + userPassword;
            pool = new V77Connection[connectionsCount];
            isReady = false;
        }
        public static V77Connection GetV77Connection(String ip)
        {
            V77Connection v77cn = null;
            if (!isReady)
            {
                fillPool(ip);
            }
            if (pool.Length > 0)
            {
                v77cn = pool[0];
                // Процесс 1с уже запущен теперь надо подключиться.
                v77cn.Open();
            }
            return v77cn;
        }
    }
}
