using System;
using System.Data;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace Nskd.Spreadsheet
{
    public class XReference
    {
        public Int32 RowIndex;
        public Int32 ColumnIndex;
        public Int32 RowCount;
        public Int32 ColumnCount;
        public XReference()
        {
            RowIndex = 0;
            ColumnIndex = 0;
            RowCount = 0;
            ColumnCount = 0;
        }
        public XReference(String xlsReference)
            : this()
        {
            if (!String.IsNullOrEmpty(xlsReference))
            {
                string[] cellReferences = xlsReference.Split(':');
                Int32[] cis;
                if (cellReferences.Length > 0)
                {
                    cis = parseXlsCellReference(cellReferences[0]);
                    RowIndex = cis[0] - 1;
                    ColumnIndex = cis[1] - 1;
                    RowCount = 1;
                    ColumnCount = 1;
                }
                if (cellReferences.Length > 1)
                {
                    cis = parseXlsCellReference(cellReferences[1]);
                    int ri = cis[0] - 1;
                    int ci = cis[1] - 1;
                    RowCount = Math.Abs(ri - RowIndex) + 1;
                    RowIndex = Math.Min(RowIndex, ri);
                    ColumnCount = Math.Abs(ci - ColumnIndex) + 1;
                    ColumnIndex = Math.Min(ColumnIndex, ci);
                }
            }
        }
        private Int32[] parseXlsCellReference(string xlsCellReference)
        {
            Int32 ri = 0;
            Int32 ci = 0;
            if (!String.IsNullOrEmpty(xlsCellReference))
            {
                Int32 i = 0;
                for (; i < xlsCellReference.Length; i++)
                {
                    Char c = xlsCellReference[i];
                    if ((c >= 'A') && (c <= 'Z'))
                    {
                        ci = (ci * 26) + ((c - 'A') + 1);
                    }
                    else if ((c >= 'a') && (c <= 'z'))
                    {
                        ci = (ci * 26) + ((c - 'a') + 1);
                    }
                    else break;
                }
                for (; i < xlsCellReference.Length; i++)
                {
                    Char c = xlsCellReference[i];
                    if ((c >= '0') && (c <= '9'))
                    {
                        ri = (ri * 10) + (c - '0');
                    }
                    else break;
                }
            }
            return (new Int32[] { ri, ci });
        }
    }
    public class XDimention : XReference
    {
        public XDimention() : base() { }
        public XDimention(String xlsReference) : base(xlsReference) { }
    }
    public class XMergeCell : XReference
    {
        public XMergeCell() : base() { }
        public XMergeCell(String xlsReference) : base(xlsReference) { }
    }
    public class XMergeCells
    {
        public Int32 Count;
        private XMergeCell[] xmcs;
        public XMergeCells()
        {
            Count = 0;
            xmcs = new XMergeCell[] { };
        }
        public XMergeCell this[Int32 i]
        {
            get
            {
                if ((i < 0) || (i >= Count)) throw new ArgumentOutOfRangeException();
                return xmcs[i];
            }
        }
        public void Add(XMergeCell xMergeCell)
        {
            Array.Resize<XMergeCell>(ref xmcs, Count + 1);
            xmcs[Count++] = xMergeCell;
        }
    }
    public class XWorksheet
    {
        public string Id;
        public string Name;
        public bool SummaryBelow;
        // моё представление листа
        // 1. таблица значений для каждой ячейки
        public DataTable CellValuesTable;
        // 2. таблица типов данных для каждой ячейки
        //public DataTable CellTypes; // не используется так как в таблице значений в каждой ячейке хранятся объект у которого есть свойство Type
        // 3. таблица стилей для каждой ячейки в html виде style="..."
        //public DataTable CellStyles;
        // 4. таблица с одним столбцом который содержит уровнь вложенности строки
        public DataTable OutlinePropertiesTable;
        public XMergeCells MergeCells;
        private DataTable cellOuterHtml;
        public XWorksheet()
        {
            Id = "";
            Name = "";
            SummaryBelow = false;
            CellValuesTable = new DataTable();
            //CellTypes = new DataTable();
            //CellStyles = new DataTable();
            OutlinePropertiesTable = new DataTable();
            MergeCells = new XMergeCells();
        }
        public XWorksheet(XDimention xd)
            : this()
        {
            int rowCount = xd.RowCount;
            int columnCount = xd.ColumnCount;
            // перенести в XWorksheet
            OutlinePropertiesTable.Columns.Add("", typeof(int));
            for (int i = 0; i < columnCount; i++)
            {
                CellValuesTable.Columns.Add("", typeof(object));
            }
            for (int i = 0; i < rowCount; i++)
            {
                OutlinePropertiesTable.Rows.Add(OutlinePropertiesTable.NewRow());
                CellValuesTable.Rows.Add(CellValuesTable.NewRow());
            }
        }
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder();
            cellOuterHtml = new DataTable();
            for (int i = 0; i < CellValuesTable.Columns.Count; i++)
            {
                cellOuterHtml.Columns.Add("", typeof(string));
            }
            for (int i = 0; i < CellValuesTable.Rows.Count; i++)
            {
                cellOuterHtml.Rows.Add(cellOuterHtml.NewRow());
            }
            if (this.CellValuesTable != null)
            {
                int maxLevel = 0;
                foreach (DataRow row in this.OutlinePropertiesTable.Rows)
                {
                    maxLevel = Math.Max(maxLevel, (int)row[0]);
                }
                html.Append("<div class='Nskd_XSheet' onclick='Nskd.Spreadsheet.XSheet.Onclick()'>");
                html.Append("<div><span>" + this.Name + "</span></div>");
                RenderRowsIntoCellOuterHtml();
                ApplyMerge();
                html.Append("<table>");
                html.Append(RenderRowsFromCellOuterHtml(maxLevel));
                html.Append("</table>");
                html.Append("</div>");
            }
            return html.ToString();
        }
        private void RenderRowsIntoCellOuterHtml()
        {
            for (int ri = 0; ri < this.CellValuesTable.Rows.Count; ri++)
            {
                DataRow row = CellValuesTable.Rows[ri];
                for (int ci = 0; ci < this.CellValuesTable.Columns.Count; ci++)
                {
                    cellOuterHtml.Rows[ri][ci] = RenderCell(row[ci]);
                }
            }
        }
        private string RenderCell(object v)
        {
            StringBuilder html = new StringBuilder();
            switch (v.GetType().ToString())
            {
                case "System.Int32":
                    html.Append("<td class='tar'>");
                    html.AppendFormat("{0:#,##0}", v);
                    html.Append("</td>");
                    break;
                case "System.Double":
                    html.Append("<td class='tar'>");
                    html.AppendFormat("{0:#,##0.00}", v);
                    html.Append("</td>");
                    break;
                case "System.DateTime":
                    html.Append("<td class='tar'>");
                    html.AppendFormat("{0:yyyy-MM-dd}", v);
                    html.Append("</td>");
                    break;
                case "System.DBNull":
                    html.Append("<td></td>");
                    break;
                default:
                    html.Append("<td class='tal'>");
                    html.Append(v);
                    html.Append("</td>");
                    break;
            }
            return html.ToString();
        }
        private void ApplyMerge()
        {
            for (int i = 0; i < MergeCells.Count; i++)
            {
                XMergeCell xmc = MergeCells[i];
                string temp = (string)cellOuterHtml.Rows[xmc.RowIndex][xmc.ColumnIndex];
                temp = temp.Replace("<td", "<td colspan='" + xmc.ColumnCount.ToString() + "' rowspan='" + "1" + "'");
                for (int ri = xmc.RowIndex; ri < (xmc.RowIndex + 1); ri++)
                {
                    for (int ci = xmc.ColumnIndex; ci < (xmc.ColumnIndex + xmc.ColumnCount); ci++)
                    {
                        cellOuterHtml.Rows[ri][ci] = "";
                    }
                }
                cellOuterHtml.Rows[xmc.RowIndex][xmc.ColumnIndex] = temp;
            }
        }
        private string RenderRowsFromCellOuterHtml(int maxLevel)
        {
            StringBuilder html = new StringBuilder();
            for (int startRowIndex = 0; startRowIndex < cellOuterHtml.Rows.Count; startRowIndex++)
            {
                DataRow row = cellOuterHtml.Rows[startRowIndex];
                int rowLevel = (int)OutlinePropertiesTable.Rows[startRowIndex][0];
                int nextRowLevel = 0;
                if ((startRowIndex + 1) < cellOuterHtml.Rows.Count)
                {
                    nextRowLevel = (int)OutlinePropertiesTable.Rows[startRowIndex + 1][0];
                }
                bool isSummary = (rowLevel < nextRowLevel);
                bool isDetail = (rowLevel > 0);
                html.Append((isDetail) ? "<tr style='display: none'>" : "<tr>");
                // скрытая ячейка с уровнем вложенности
                html.Append("<td class='dn'>" + rowLevel.ToString() + "</td>");
                // ячейки для управления подчинёнными строками
                for (int i = 0; i < maxLevel; i++)
                {
                    html.Append("<td class='expand'>");
                    html.Append(((isSummary) && (i == rowLevel)) ? "+" : "&nbsp;"); // или "-"
                    html.Append("</td>");
                }
                // все остальные ячейки с данными
                for (int ci = 0; ci < cellOuterHtml.Columns.Count; ci++)
                {
                    html.Append(cellOuterHtml.Rows[startRowIndex][ci]);
                }
                html.Append("</tr>");
            }
            return html.ToString();
        }
    }
    public class XWorksheets
    {
        public Int32 Count;
        private XWorksheet[] xwss;
        public XWorksheets()
        {
            Count = 0;
            xwss = new XWorksheet[] { };
        }
        public XWorksheet this[Int32 i]
        {
            get
            {
                if ((i < 0) || (i >= Count)) throw new ArgumentOutOfRangeException();
                return xwss[i];
            }
        }
        public void Add(XWorksheet xWorksheet)
        {
            Array.Resize<XWorksheet>(ref xwss, Count + 1);
            xwss[Count++] = xWorksheet;
        }
    }
    public class XDocument
    {
        public XWorksheets XWorksheets;
        // Ссылки на части пакета.
        private SharedStringTable sharedStringTable;
        private Stylesheet stylesheet;

        public XDocument()
        {
            XWorksheets = new XWorksheets();
            sharedStringTable = null;
            stylesheet = null;
        }
        public void LoadXlsx(string path)
        {
            // Буфер - чтобы не держать файл открытым.
            MemoryStream memoryStream = new MemoryStream();
            using (Stream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                stream.CopyTo(memoryStream);
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(memoryStream, false))
            {
                WorkbookPart workbookPart = package.WorkbookPart;
                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                this.sharedStringTable = sharedStringTablePart.SharedStringTable;
                WorkbookStylesPart workbookStylesPart = workbookPart.WorkbookStylesPart;
                this.stylesheet = workbookStylesPart.Stylesheet;
                Workbook workbook = workbookPart.Workbook;
                Sheets sheets = workbook.Sheets;
                foreach (Sheet sheet in sheets.Descendants<Sheet>())
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    // размеры
                    XDimention xd = new XDimention(worksheet.SheetDimension.Reference);
                    // добавляем новый лист в представление Net
                    XWorksheet xWorksheet = new XWorksheet(xd);
                    XWorksheets.Add(xWorksheet);
                    xWorksheet.Name = sheet.Name;
                    // разбор свойств листа
                    parseSheetProperties(worksheet, xWorksheet);
                    // разбор данных
                    parseSheetData(worksheet, xWorksheet);
                    // разбор объединений ячеек
                    parseMegreCells(worksheet, xWorksheet);
                }
            }
            memoryStream.Dispose();
        }
        public string ToHtml()
        {
            StringBuilder html = new StringBuilder();
            for (int i = 0; i < XWorksheets.Count; i++)
            {
                html.Append("<br />");
                html.Append("<div>");
                html.Append(XWorksheets[i].ToHtml());
                html.Append("</div>");
                html.Append("<br />");
            }
            return html.ToString();
        }

        private void parseSheetProperties(Worksheet worksheet, XWorksheet xWorksheet)
        {
            if (worksheet.SheetProperties != null)
            {
                if (worksheet.SheetProperties.OutlineProperties != null)
                {
                    xWorksheet.SummaryBelow = worksheet.SheetProperties.OutlineProperties.SummaryBelow;
                }
            }
        }
        private void parseSheetData(Worksheet worksheet, XWorksheet xWorksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            foreach (Row row in sheetData.Descendants<Row>())
            {
                int dtRowIndex = (int)(uint)row.RowIndex - 1;
                // группировка строк
                DataRow opRow = xWorksheet.OutlinePropertiesTable.Rows[dtRowIndex];
                opRow[0] = (row.OutlineLevel == null) ? 0 : (int)row.OutlineLevel;
                // данные
                DataRow cvRow = xWorksheet.CellValuesTable.Rows[dtRowIndex];
                foreach (Cell cell in row.Descendants<Cell>())
                {
                    XReference xr = new XReference(cell.CellReference);
                    cvRow[xr.ColumnIndex] = getCellValue(cell);
                }
            }
        }
        private void parseMegreCells(Worksheet worksheet, XWorksheet xWorksheet)
        {
            foreach (MergeCells mergeCells in worksheet.Descendants<MergeCells>())
            {
                foreach (MergeCell mergeCell in mergeCells.Descendants<MergeCell>())
                {
                    XMergeCell xmc = new XMergeCell(mergeCell.Reference);
                    xWorksheet.MergeCells.Add(xmc);
                }
            }
        }
        private object getCellValue(Cell cell)
        {
            object v = null;
            if (cell.CellValue != null) // Есть tag x:v
            {
                double cellValue; // или индекс или значение
                if (double.TryParse(cell.CellValue.Text, out cellValue))
                {
                    var dataType = cell.DataType;
                    if (dataType != null) // Есть атрибут t.
                    {
                        // всегда стока или из sharedStringTable или из InlineString
                        v = getCellValueForAttrT(cell, cellValue);
                    }
                    else // Нет атрибута t.
                    {
                        // Попробуем разобрать атрибут s.
                        v = getCellValueForAttrS(cell, cellValue);
                    }
                }
            }
            return v;
        }
        private string getCellValueForAttrT(Cell cell, double cellValue)
        {
            // Точно есть cell.CellValue и cell.Datatype. Это уже проверено до вызова.
            // всегда стока или из sharedStringTable или из InlineString
            string v = "";
            CellValues dataType = (CellValues)cell.DataType;
            switch (dataType)
            {
                case CellValues.SharedString:
                    int ssti = (int)cellValue; // индекс
                    v = this.sharedStringTable.ChildElements[ssti].FirstChild.InnerText;
                    break;
                case CellValues.InlineString:
                    v = cell.InlineString.Text.Text;
                    break;
                default:
                    v = dataType.ToString() + ": " + cell.CellValue.Text;
                    break;
            }
            return v;
        }
        private object getCellValueForAttrS(Cell cell, double cellValue)
        {
            object v = null;

            var styleIndex = cell.StyleIndex; // атрибут s
            if (styleIndex != null)
            {
                int si;
                if (int.TryParse(styleIndex.Value.ToString(), out si))
                {
                    CellFormats cfs = this.stylesheet.CellFormats;
                    var en = cfs.Descendants<CellFormat>().GetEnumerator();
                    while (si-- >= 0) en.MoveNext();
                    CellFormat cf = en.Current;
                    switch (cf.NumberFormatId.Value)
                    {
                        case 0:
                        case 1:
                            int i = (int)cellValue;
                            v = i; // String.Format("{0:0}", i);
                            break;
                        case 4:
                        case 167:
                            v = cellValue; // String.Format("{0:#,##0.00}", cellValue);
                            break;
                        case 14:
                        case 164:
                        case 165:
                        case 166:
                            int days = (int)cellValue;
                            DateTime date = (new DateTime(1900, 1, 1)).AddDays(days - 2);
                            v = date; // String.Format("{0:yyyy-MM-dd}", date);
                            break;
                        default:
                            v = cf.NumberFormatId.Value.ToString() + ": " + cell.CellValue.Text;
                            break;
                    }
                }
                else
                {
                    v = cellValue;
                }
            }
            return v;
        }
    }
}
