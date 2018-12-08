using System;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Nskd.V83
{
    public class COMОбъект : IDisposable
    {
        private const BindingFlags INVOKE_METHOD = BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static;
        private const BindingFlags GET_PROPERTY = BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Static;
        private const BindingFlags SET_PROPERTY = BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.Static;

        public Object ComObject { get; private set; }
        private Type oType;

        public COMОбъект() { ComObject = null; oType = null; }
        public COMОбъект(Object comObject)
        {
            if (comObject != null && comObject.GetType().Name == "__ComObject")
            {
                ComObject = comObject;
                oType = ComObject.GetType();
            }
            else { ComObject = null; oType = null; }
        }

        public Object GetProperty(String name)
        {
            return oType?.InvokeMember(name, GET_PROPERTY, null, ComObject, null);
        }

        public void SetProperty(String name, Object value)
        {
            if (value is COMОбъект) value = ((COMОбъект)value).ComObject;
            oType?.InvokeMember(name, SET_PROPERTY, null, ComObject, new Object[] { value });
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

namespace Nskd.V83
{
    /// <summary>
    ///     Глобальный контекст для внешнего соединения
    /// </summary>
    public class GlobalContext : COMОбъект
    {
        private static Object GetContext(String cnString)
        {
            var connector = new global::V83.COMConnector();
            var connection = connector.Connect(cnString);
            return connection;
        }

        public GlobalContext(String cnString) : base(GetContext(cnString)) { }

        public ПрикладныеОбъекты.СправочникиМенеджер Справочники
        {
            get
            {
                return new ПрикладныеОбъекты.СправочникиМенеджер(GetProperty("Справочники"));
            }
        }
        public ПрикладныеОбъекты.ПеречисленияМенеджер Перечисления
        {
            get { return new ПрикладныеОбъекты.ПеречисленияМенеджер(GetProperty("Перечисления")); }
        }
        public ПрикладныеОбъекты.ПланыВидовХарактеристикМенеджер ПланыВидовХарактеристик
        {
            get { return new ПрикладныеОбъекты.ПланыВидовХарактеристикМенеджер(GetProperty("ПланыВидовХарактеристик")); }
        }
        public ПрикладныеОбъекты.РегистрыСведенийМенеджер РегистрыСведений
        {
            get
            {
                return new ПрикладныеОбъекты.РегистрыСведенийМенеджер(GetProperty("РегистрыСведений"));
            }
        }
        public РаботаСЗапросами.Запрос Запрос
        {
            get { return new РаботаСЗапросами.Запрос(InvokeMethod("NewObject", "Запрос")); }
        }

        /// <summary>
        ///     NewObject(Имя)
        ///     Создает объект, для которого предусмотрен конструктор, и возвращает ссылку на него.
        ///     Последующие параметры метода те же, что у конструктора объекта, 
        ///     имя которого указано в качестве значения первого параметра.
        /// </summary>
        /// <param name="Имя">
        ///     Имя (обязательный)
        ///     Тип: Строка.
        ///     Имя объекта, объявленного в конфигураторе.
        /// </param>
        public COMОбъект NewObject(String Имя)
        {
            return new COMОбъект(InvokeMethod("NewObject", Имя));
        }

        // Процедуры и функции работы с XML
        public String XMLСтрока(Object Значение)
        {
            return (String)InvokeMethod("XMLСтрока", Значение);
        }

        // Процедуры и функции работы с информационной базой
        public void ЗафиксироватьТранзакцию()
        {
            InvokeMethod("ЗафиксироватьТранзакцию");
        }
        public void НачатьТранзакцию()
        {
            InvokeMethod("НачатьТранзакцию");
        }
        public void ОтменитьТранзакцию()
        {
            InvokeMethod("ОтменитьТранзакцию");
        }

        // Функции преобразования значений
        public String Строка(Object Значение)
        {
            return (String)InvokeMethod("String", Значение);
        }
    }

    public class V83Document : COMОбъект
    {
        public V83Document(Object comObject) : base(comObject) { }

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
        public Double НайтиПоНомеру(String num, DateTime date, String kind = "")
        {
            return (Double)InvokeMethod("НайтиПоНомеру", new Object[] { num, date, kind });
        }

        /// <summary>
        /// ТекущийДокумент()
        /// Метод ТекущийДокумент() возвращает значение позиционированного текущего документа (в целом, как объекта). 
        /// Данный метод применяется, например, если нужно документ передать как параметр в вызове какого-либо метода или присвоить какому-либо реквизиту.
        /// </summary>
        /// <returns>Значение текущего документа.</returns>
        public COMОбъект ТекущийДокумент()
        {
            return new COMОбъект(InvokeMethod("ТекущийЭлемент"));
        }
    }

    public static class OcStoredProcedures
    {
        private static Object thisLock = new Object();
        public static ResponsePackage Exec1(RequestPackage rqp)
        {
            ResponsePackage rsp = new ResponsePackage();
            DataTable dt = null;
            lock (thisLock)
            {
                // Запускаем 1с или подключаемся к уже запущенному.
                // Root COM object 1c 'System'.
                GlobalContext sys = new GlobalContext(@"Srvr=""srv-82:1741""; Ref=""BUH"";Usr=""Соколов Евгений COM0"";Pwd=""yNFxfrvqxP"";");
                if (sys != null)
                {
                    try
                    {
                        //dt = GetPartnerList(sys);
                        //dt = GetPaymentList(sys);
                        /*
                        switch (rqp.Command)
                        {
                            case "Добавить":
                                //code = Agrs.F0Add(v77, rqp);
                                break;
                            case "Обновить":
                                //code = Agrs.F0Update(v77, rqp);
                                break;
                            case "Удалить":
                                //code = Agrs.F0Delete(v77, rqp);
                                break;
                            case "Docs1c/F0/Save":
                                //Docs1c.Save(v77, rqp);
                                break;
                            default:
                                break;
                        }
                        */
                    }
                    catch (Exception e) { Console.WriteLine(e.ToString()); }
                    finally
                    {
                        //if (sys != null) { sys.Release(); }
                        //GC.Collect();
                    }
                }
            }
            rsp.Data = new DataSet();
            if (dt != null) { rsp.Data.Tables.Add(dt); }
            return rsp;
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

namespace Nskd.V83.ОбщиеОбъекты
{
    // Работа с объектами метаданных
    public class ОбъектыМетаданных
    {
        public class Перечисление : COMОбъект
        {
            public Перечисление(Object comObject) : base(comObject) { }

            public String Имя { get { return (String)GetProperty("Имя"); } }
        }
    }

    // Управление блокировкой данных
    public class БлокировкаДанных : COMОбъект
    {
        public БлокировкаДанных(Object comObject) : base(comObject) { }
        public ЭлементБлокировкиДанных Добавить(String ПространствоБлокировки)
        {
            return new ЭлементБлокировкиДанных(InvokeMethod("Добавить", ПространствоБлокировки));
        }
        public void Заблокировать()
        {
            InvokeMethod("Заблокировать");
        }
    }
    public class ЭлементБлокировкиДанных : COMОбъект
    {
        public ЭлементБлокировкиДанных(Object comObject) : base(comObject) { }
    }
}

namespace Nskd.V83.ПрикладныеОбъекты
{
    // Документы
    public class ДокументОбъект : COMОбъект
    {
        public ДокументОбъект(Object comObject) : base(comObject)
        {
        }

        public Boolean ПометкаУдаления
        {
            get { return (Boolean)GetProperty("ПометкаУдаления"); }
            set { SetProperty("ПометкаУдаления", value); }
        }

        public void Записать()
        {
            InvokeMethod("Записать");
        }
        public void Удалить()
        {
            InvokeMethod("Удалить");
        }
        public void УстановитьПометкуУдаления(Boolean ПометкаУдаления)
        {
            InvokeMethod("УстановитьПометкуУдаления", ПометкаУдаления);
        }
    }
    public class ДокументСсылка : COMОбъект
    {
        public ДокументСсылка(Object comObject) : base(comObject)
        {
        }

        public String Номер { get { return (String)GetProperty("Номер"); } }

        public ДокументОбъект ПолучитьОбъект()
        {
            return new ДокументОбъект(InvokeMethod("ПолучитьОбъект"));
        }
        public Boolean Пустая()
        {
            return (Boolean)InvokeMethod("Пустая");
        }
    }

    // Справочники
    public class СправочникиМенеджер : COMОбъект
    {
        public СправочникиМенеджер(Object comObject) : base(comObject) { }

        public СправочникМенеджер ВидыНоменклатуры
        {
            get
            {
                return new СправочникМенеджер(GetProperty("ВидыНоменклатуры"));
            }
        }
        public СправочникМенеджер Номенклатура
        {
            get
            {
                return new СправочникМенеджер(GetProperty("Номенклатура"));
            }
        }
        public СправочникМенеджер Производители
        {
            get
            {
                return new СправочникМенеджер(GetProperty("Производители"));
            }
        }
        public СправочникМенеджер УпаковкиЕдиницыИзмерения
        {
            get
            {
                return new СправочникМенеджер(GetProperty("УпаковкиЕдиницыИзмерения"));
            }
        }
        public СправочникМенеджер ЗначенияСвойствОбъектов
        {
            get
            {
                return new СправочникМенеджер(GetProperty("ЗначенияСвойствОбъектов"));
            }
        }
    }
    public class СправочникМенеджер : COMОбъект
    {
        public СправочникМенеджер(Object comObject) : base(comObject)
        {
        }
        public СправочникВыборка Выбрать(СправочникСсылка Родитель = null, СправочникСсылка Владелец = null, Object Отбор = null, String Порядок = null)
        {
            return new СправочникВыборка(InvokeMethod("Выбрать", new Object[] { Родитель?.ComObject, Владелец?.ComObject, Отбор, Порядок }));
        }
        public СправочникВыборка ВыбратьИерархически(СправочникСсылка Родитель = null, СправочникСсылка Владелец = null, Object Отбор = null, String Порядок = null)
        {
            return new СправочникВыборка(InvokeMethod("ВыбратьИерархически", new Object[] { Родитель?.ComObject, Владелец?.ComObject, Отбор, Порядок }));
        }
        public СправочникСсылка НайтиПоНаименованию(String descr, Boolean exactly = false, СправочникСсылка Родитель = null, СправочникСсылка Владелец = null)
        {
            return new СправочникСсылка(InvokeMethod("НайтиПоНаименованию", new Object[] { descr, exactly, Родитель?.ComObject, Владелец?.ComObject }));
        }
        public СправочникСсылка НайтиПоРеквизиту(String ИмяРеквизита, Object ЗначениеРеквизита, СправочникСсылка Родитель = null, СправочникСсылка Владелец = null)
        {
            return new СправочникСсылка(InvokeMethod("НайтиПоРеквизиту", new Object[] { ИмяРеквизита, ЗначениеРеквизита, Родитель?.ComObject, Владелец?.ComObject }));
        }
        public СправочникСсылка ПустаяСсылка()
        {
            return new СправочникСсылка(InvokeMethod("ПустаяСсылка"));
        }
        public СправочникОбъект СоздатьГруппу()
        {
            return new СправочникОбъект(InvokeMethod("СоздатьГруппу"));
        }
        public СправочникОбъект СоздатьЭлемент()
        {
            return new СправочникОбъект(InvokeMethod("СоздатьЭлемент"));
        }
    }
    public class СправочникСсылка : COMОбъект
    {
        public СправочникСсылка(Object comObject) : base(comObject)
        {
        }
        public static implicit operator СправочникСсылка(ПланВидовХарактеристикСсылка ссылка)
        {
            return new СправочникСсылка(ссылка.ComObject);
        }


        public String Наименование { get { return (String)GetProperty("Наименование"); } }
        public СправочникСсылка Родитель { get { return new СправочникСсылка(GetProperty("Родитель")); } }
        public Boolean ЭтоГруппа { get { return (Boolean)GetProperty("ЭтоГруппа"); } }

        public СправочникОбъект ПолучитьОбъект()
        {
            return new СправочникОбъект(InvokeMethod("ПолучитьОбъект"));
        }
        public Boolean Пустая()
        {
            return (Boolean)InvokeMethod("Пустая");
        }
    }
    public class СправочникОбъект : COMОбъект
    {
        public СправочникОбъект(Object comObject) : base(comObject)
        {
        }

        public СправочникСсылка Владелец
        {
            get { return new СправочникСсылка(GetProperty("Владелец")); }
            set { SetProperty("Владелец", value); }
        }
        public String Наименование
        {
            get { return GetProperty("Наименование") as String; }
            set { SetProperty("Наименование", value); }
        }
        public Boolean ПометкаУдаления
        {
            get { return (Boolean)GetProperty("ПометкаУдаления"); }
            set { SetProperty("ПометкаУдаления", value); }
        }
        public СправочникСсылка Родитель
        {
            get { return new СправочникСсылка(GetProperty("Родитель")); }
            set { SetProperty("Родитель", value); }
        }
        public СправочникСсылка Ссылка
        {
            get { return new СправочникСсылка(GetProperty("Ссылка")); }
        }

        public void Записать()
        {
            InvokeMethod("Записать");
        }
        public void Удалить()
        {
            InvokeMethod("Удалить");
        }
        public void УстановитьПометкуУдаления(Boolean ПометкаУдаления, Boolean ВключаяПодчиненные = true)
        {
            InvokeMethod("УстановитьПометкуУдаления", new Object[] { ПометкаУдаления, ВключаяПодчиненные });
        }
    }
    public class СправочникВыборка : COMОбъект
    {
        public СправочникВыборка(Object comObject) : base(comObject) { }

        public String Наименование
        {
            get { return GetProperty("Наименование") as String; }
        }
        public СправочникСсылка Родитель
        {
            get { return new СправочникСсылка(GetProperty("Родитель")); }
        }
        public СправочникСсылка Ссылка
        {
            get { return new СправочникСсылка(GetProperty("Ссылка")); }
        }
        public Boolean ЭтоГруппа
        {
            get { return (Boolean)GetProperty("ЭтоГруппа"); }
        }

        public СправочникОбъект ПолучитьОбъект()
        {
            return new СправочникОбъект(InvokeMethod("ПолучитьОбъект"));
        }
        public Boolean Следующий()
        {
            return (Boolean)InvokeMethod("Следующий");
        }
    }
    public class СправочникСписок : COMОбъект
    {
        public СправочникСписок(Object comObject) : base(comObject) { }
    }

    // Перечисления
    public class ПеречисленияМенеджер : COMОбъект
    {
        public ПеречисленияМенеджер(Object comObject) : base(comObject) { }

        public ПеречислениеМенеджер ВариантыИспользованияХарактеристикНоменклатуры { get { return new ПеречислениеМенеджер(this.GetProperty("ВариантыИспользованияХарактеристикНоменклатуры")); } }
        public ПеречислениеМенеджер ВариантыОформленияПродажи { get { return new ПеречислениеМенеджер(this.GetProperty("ВариантыОформленияПродажи")); } }
        public ПеречислениеМенеджер ТипыНоменклатуры { get { return new ПеречислениеМенеджер(this.GetProperty("ТипыНоменклатуры")); } }
        public ПеречислениеМенеджер СтавкиНДС { get { return new ПеречислениеМенеджер(this.GetProperty("СтавкиНДС")); } }
    }
    public class ПеречислениеМенеджер : COMОбъект
    {
        public ПеречислениеМенеджер(Object comObject) : base(comObject) { }

        public Double Индекс(ПеречисленияСсылка ЗначениеПеречисления)
        {
            return (Double)InvokeMethod("Индекс", ЗначениеПеречисления);
        }
    }
    public class ПеречисленияСсылка : COMОбъект
    {
        public ПеречисленияСсылка(Object comObject) : base(comObject) { }

        public ОбщиеОбъекты.ОбъектыМетаданных.Перечисление Метаданные()
        {
            return new ОбщиеОбъекты.ОбъектыМетаданных.Перечисление(InvokeMethod("Метаданные"));
        }
        public Boolean Пустая()
        {
            return (Boolean)InvokeMethod("Пустая");
        }
    }

    // ПланыВидовХарактеристик
    public class ПланыВидовХарактеристикМенеджер : COMОбъект
    {
        public ПланыВидовХарактеристикМенеджер(Object comObject) : base(comObject)
        {
        }

        public ПланВидовХарактеристикМенеджер ДополнительныеРеквизитыИСведения
        {
            get { return new ПланВидовХарактеристикМенеджер(GetProperty("ДополнительныеРеквизитыИСведения")); }
        }
    }
    public class ПланВидовХарактеристикМенеджер : COMОбъект
    {
        public ПланВидовХарактеристикМенеджер(Object comObject) : base(comObject)
        {
        }

        public ПланВидовХарактеристикСсылка НайтиПоНаименованию(String Наименование, Boolean ТочноеСоответствие = false, ПланВидовХарактеристикСсылка Родитель = null)
        {
            return new ПланВидовХарактеристикСсылка(InvokeMethod("НайтиПоНаименованию", new Object[] { Наименование, ТочноеСоответствие, Родитель }));
        }
    }
    public class ПланВидовХарактеристикСсылка : COMОбъект
    {
        public ПланВидовХарактеристикСсылка(Object comObject) : base(comObject)
        {
        }

        public String Наименование
        {
            get { return GetProperty("Наименование") as String; }
        }
    }

    // Регистры сведений
    public class РегистрыСведенийМенеджер : COMОбъект
    {
        public РегистрыСведенийМенеджер(Object comObject) : base(comObject) { }

        public РегистрСведенийМенеджер ШтрихкодыНоменклатуры
        {
            get
            {
                return new РегистрСведенийМенеджер(GetProperty("ШтрихкодыНоменклатуры"));
            }
        }
    }
    public class РегистрСведенийМенеджер: COMОбъект
    {
        public РегистрСведенийМенеджер(Object comObject) : base(comObject) { }

        public РегистрСведенийМенеджерЗаписи СоздатьМенеджерЗаписи()
        {
            return new РегистрСведенийМенеджерЗаписи(InvokeMethod("СоздатьМенеджерЗаписи"));
        }
        public РегистрСведенийНаборЗаписей СоздатьНаборЗаписей()
        {
            return new РегистрСведенийНаборЗаписей(InvokeMethod("СоздатьНаборЗаписей"));
        }
    }
    public class РегистрСведенийМенеджерЗаписи : COMОбъект
    {
        public РегистрСведенийМенеджерЗаписи(Object comObject) : base(comObject) { }

        public void Записать(Boolean Замещать)
        {
            InvokeMethod("Записать", Замещать);
        }
    }
    public class РегистрСведенийНаборЗаписей : COMОбъект
    {
        public РегистрСведенийНаборЗаписей(Object comObject) : base(comObject) { }


        public РегистрСведенийЗапись Добавить()
        {
            return new РегистрСведенийЗапись(InvokeMethod("Добавить"));
        }
        public Int32 Количество()
        {
            return (Int32)InvokeMethod("Количество");
        }
        public РегистрСведенийЗапись Получить(Int32 Индекс)
        {
            return new РегистрСведенийЗапись(InvokeMethod("Получить", Индекс));
        }
        public void Прочитать()
        {
            InvokeMethod("Прочитать");
        }
    }
    public class РегистрСведенийЗапись : COMОбъект
    {
        public РегистрСведенийЗапись(Object comObject) : base(comObject) { }
    }
}

namespace Nskd.V83.ПрикладныеОбъекты.УниверсальныеОбъекты
{
    public class ТабличнаяЧасть : COMОбъект
    {
        public ТабличнаяЧасть(Object comObject) : base(comObject)
        {
        }

        public СтрокаТабличнойЧасти Добавить()
        {
            return new СтрокаТабличнойЧасти(InvokeMethod("Добавить"));
        }
        public Int32 Количество()
        {
            return (Int32)InvokeMethod("Количество");
        }
        public СтрокаТабличнойЧасти Найти(Object Значение, String Колонки = "")
        {
            return new СтрокаТабличнойЧасти(InvokeMethod("Найти", new Object[] { Значение, Колонки }));
        }
        public СтрокаТабличнойЧасти Получить(Int32 Индекс)
        {
            return new СтрокаТабличнойЧасти(InvokeMethod("Получить", Индекс));
        }
        public void Удалить(Int32 Индекс)
        {
            InvokeMethod("Удалить", Индекс);
        }

        public СтрокаТабличнойЧасти this[Int32 index]
        {
            get
            {
                return new СтрокаТабличнойЧасти(InvokeMethod("Получить", index));
            }
        }
    }
    public class СтрокаТабличнойЧасти : COMОбъект
    {
        public СтрокаТабличнойЧасти(Object comObject) : base(comObject)
        {
        }

        public Int32 НомерСтроки
        {
            get
            {
                return (Int32)GetProperty("НомерСтроки");
            }
        }
    }
}

namespace Nskd.V83.РаботаСЗапросами
{
    // Выполнение и работа с запросами во встроенном языке
    public class Запрос : COMОбъект
    {
        public Запрос(Object comObject) : base(comObject) { }

        public String Текст
        {
            get { return GetProperty("Текст") as String; }
            set { SetProperty("Текст", value); }
        }

        public РезультатЗапроса Выполнить()
        {
            Object comObject = InvokeMethod("Выполнить");
            РезультатЗапроса r = new РезультатЗапроса(comObject);
            return r;
        }
        public void УстановитьПараметр(String Имя, Object Значение)
        {
            InvokeMethod("УстановитьПараметр", new Object[] { Имя, Значение });
        }
    }

    public class РезультатЗапроса : COMОбъект
    {
        public РезультатЗапроса(Object comObject) : base(comObject) { }

        /// <summary>
        ///     Выбрать(<ТипОбхода>, <Группировки>, <ГруппировкиДляЗначенийГруппировок>) 
        /// </summary>
        /// <returns></returns>
        public ВыборкаИзРезультатаЗапроса Выбрать()
        {
            return new ВыборкаИзРезультатаЗапроса(InvokeMethod("Выбрать"));
        }
    }

    public class ВыборкаИзРезультатаЗапроса : COMОбъект
    {
        public ВыборкаИзРезультатаЗапроса(Object comObject) : base(comObject) { }

        public Int32 Количество()
        {
            return (Int32)InvokeMethod("Количество");
        }

        public Boolean Следующий()
        {
            return (Boolean)InvokeMethod("Следующий");
        }
    }

}
