using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataExporting
{
    public class ExcelWriter : DataFileWriter
    {
        #region Excel COM Objects

        private readonly object NoValue = Missing.Value;

        private Excel.Application _application;

        private Excel.Workbooks _workbooks;

        private Excel.Workbook _workbook;

        private Excel.Sheets _worksheets;

        private Excel.Worksheet _currentWorksheet;

        #endregion

        #region Static Properties

        public static readonly TableStyle DefaultTableStyle = new TableStyle()
        {
            FieldNameStyle = NoStyle.Instance,
            BodyStyle = NoStyle.Instance
        };

        private static readonly Dictionary<HorizontalAlignment, Excel.XlHAlign> horizontalAlignmentMaps
                = new Dictionary<HorizontalAlignment, Excel.XlHAlign>()
                {
                    { HorizontalAlignment.Variable, Excel.XlHAlign.xlHAlignGeneral },
                    { HorizontalAlignment.Left, Excel.XlHAlign.xlHAlignLeft },
                    { HorizontalAlignment.Middle, Excel.XlHAlign.xlHAlignCenter },
                    { HorizontalAlignment.Right, Excel.XlHAlign.xlHAlignRight },
                    { HorizontalAlignment.SpacedIn, Excel.XlHAlign.xlHAlignCenterAcrossSelection },
                    { HorizontalAlignment.SpacedOut, Excel.XlHAlign.xlHAlignDistributed },
                };

        private static readonly Dictionary<VerticalAlignment, Excel.XlVAlign> verticalAlignmentMaps
            = new Dictionary<VerticalAlignment, Excel.XlVAlign>()
            {
                    { VerticalAlignment.Bottom, Excel.XlVAlign.xlVAlignBottom },
                    { VerticalAlignment.Center, Excel.XlVAlign.xlVAlignCenter },
                    { VerticalAlignment.Top, Excel.XlVAlign.xlVAlignTop },
                    { VerticalAlignment.SpacedIn, Excel.XlVAlign.xlVAlignDistributed },
                    { VerticalAlignment.SpacedOut, Excel.XlVAlign.xlVAlignJustify},
            };

        #endregion

        #region Properties

        private int _currentRow;

        private int _currentColumn;

        private int documentColumnOffset = 0;

        public int DocumentColumnOffset
        {
            get => documentColumnOffset;
            set
            {
                documentColumnOffset = value > 0 ? value : 0;
                if (_currentColumn < BaseColumn) _currentColumn = BaseColumn;
            }
        }

        private int documentRowOffset = 0;
        
        public int DocumentRowOffset
        {
            get => documentRowOffset;
            set
            {
                documentRowOffset = value > 0 ? value : 0;
                if (_currentRow < BaseRow) _currentRow = BaseRow;
            }
        }

        private int BaseRow => 1 + DocumentRowOffset;

        private int BaseColumn => 1 + DocumentColumnOffset;

        private TableStyle _dataTableStyle = DefaultTableStyle;

        public TableStyle DataTableStyle
        {
            get
            {
                if (_dataTableStyle == null) _dataTableStyle = DefaultTableStyle;
                return _dataTableStyle;
            }
            set => _dataTableStyle = value;
        }

        public WriterStyle CellStyle { get; set; } = NoStyle.Instance;

        #endregion

        #region Factory Method

        public static ExcelWriter LoadExcel(string filePath)
        {
            return new ExcelWriter(filePath);
        }

        #endregion

        #region Constructors

        public ExcelWriter()
        {
            _application = new Excel.Application();
            _workbooks = _application.Workbooks;
            _workbook = _workbooks.Add(NoValue);
            _worksheets = _workbook.Worksheets;
            _currentWorksheet = (Excel.Worksheet)_worksheets.get_Item(1);
        }

        private ExcelWriter(string filePath)
        {
            TargetPath = filePath;
            _application = new Excel.Application();
            _workbooks = _application.Workbooks;
            _workbook = _application.Workbooks.Open(filePath, ReadOnly: false, Editable: true);
            _worksheets = _workbook.Worksheets;
            _currentWorksheet = (Excel.Worksheet)_workbook.Sheets[1];

            _currentRow = _currentWorksheet.Cells.Find("*", NoValue, NoValue, NoValue, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, NoValue, NoValue).Row;

            _currentColumn = _currentWorksheet.Cells.Find("*", NoValue, NoValue, NoValue, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, NoValue, NoValue).Column;
        }

        #endregion

        #region Public Methods

        public ExcelWriter SkipRows(int count = 1)
        {
            _currentRow += (count = count > 0 ? count : 1);
            return this;
        }

        public ExcelWriter SkipColumns(int count = 1)
        {
            _currentColumn += (count = count > 0 ? count : 1);
            return this;
        }

        public ExcelWriter NewRow()
        {
            ++_currentRow;
            _currentColumn = BaseColumn;
            return this;
        }

        public ExcelWriter ResetColumn()
        {
            _currentColumn = BaseColumn;
            return this;
        }

        public ExcelWriter ResetRow()
        {
            _currentRow = BaseRow;
            return this;
        }

        public ExcelWriter AppendTable<TData>(IEnumerable<TData> source, string title = "") where TData : class
        {
            TData check = source.ToArray()[0];


            Dictionary<int, FieldData> fieldCache = new Dictionary<int, FieldData>();

            Type type = check.GetType();

            PropertyInfo[] properties = type.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly);

            int propertyCount = properties.Length;

            int dictionaryIndex = 0;

            for (int i = 0; i < propertyCount; ++i)
            {
                PropertyInfo property = properties[i];

                if (property.CanRead)
                {
                    var attributes = property.GetCustomAttributes(false);

                    FieldDisplayAttribute fieldDisplay = (FieldDisplayAttribute)attributes.FirstOrDefault(a => a.GetType() == typeof(FieldDisplayAttribute));

                    if (fieldDisplay == null) fieldCache.Add(dictionaryIndex++, new FieldData(property));
                    else if (fieldDisplay.IsIncluded) fieldCache.Add(dictionaryIndex++, new FieldData(property, (fieldDisplay.Title != string.Empty ? fieldDisplay.Title : property.Name)));
                }
            }

            propertyCount = fieldCache.Count;

            if (fieldCache.Count != 0)
            {
                int tableStartColumn = _currentColumn;

                Excel.Range tableBegin = (Excel.Range)_currentWorksheet.Cells[_currentRow, tableStartColumn];

                Excel.Range tableEnd;

                //Writes the table name
                if (title != string.Empty)
                {
                    Excel.Range tableLastColumn = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn + propertyCount - 1];

                    Excel.Range tableTitleRange = _currentWorksheet.Range[tableBegin, tableLastColumn];

                    tableTitleRange.Merge();

                    DataTableStyle.TitleStyle.ApplyStyle(tableTitleRange);

                    tableTitleRange.Cells[1, 1] = title;

                    ++_currentRow;
                }

                //Writes the Column names in table

                Excel.Range columnsBegin = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn];

                for (int i = 0; i < propertyCount; ++i)
                {
                    _currentWorksheet.Cells[_currentRow, _currentColumn++] = fieldCache[i].FieldName;
                }
                Excel.Range columnsEnd = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn - 1];
                _dataTableStyle.FieldNameStyle.ApplyStyle(_currentWorksheet.Range[columnsBegin, columnsEnd]);
                _currentColumn = tableStartColumn;
                ++_currentRow;

                int rowCount = source.Count();

                TData[] dataCopy = source.ToArray();

                int tableStartRow = _currentRow;

                Excel.Range tableBodyStart = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn];

                for (int i = 0; i < propertyCount; ++i)
                {
                    PropertyInfo property = fieldCache[i].Property;

                    int currentColumn = i + 1;

                    Excel.Range columnStart = (Excel.Range)_currentWorksheet.Cells[_currentRow, currentColumn];

                    for (int j = 0; j < rowCount; ++j)
                    {
                        _currentWorksheet.Cells[_currentRow++, _currentColumn] = property.GetValue(dataCopy[j], null).ToString();
                    }

                    _currentRow = tableStartRow;
                    ++_currentColumn;
                }

                columnsEnd.AutoFilter(1, NoValue, Excel.XlAutoFilterOperator.xlAnd, NoValue, true);

                _currentRow = tableStartRow + rowCount - 1;

                tableEnd = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn - 1];

                Excel.Range tableBody = _currentWorksheet.Range[tableBodyStart, tableEnd];

                Excel.Range table = _currentWorksheet.Range[tableBegin, tableEnd];

                _dataTableStyle.BodyStyle.ApplyStyle(tableBody);

                table.Columns.AutoFit();

                _currentColumn = tableStartColumn;

                Deallocate(tableBegin, tableEnd, table, tableBody, tableBodyStart, columnsBegin, columnsEnd);
            }
            return this;
        }

        public ExcelWriter AppendChart(IDictionary<string, double> keyValuePairs, int widthPX, int heightPX, string seriesName)
        {
            //chart.ChartType = Excel.XlChartType.xlLine;

            var yvalues = keyValuePairs.Select(kvp => kvp.Value).ToArray();

            var xvalues = keyValuePairs.Select(kvp => kvp.Key).ToArray();

            Excel.Range documentStart = (Excel.Range)_currentWorksheet.Cells[1, 1];

            Excel.Range documentEnd = (Excel.Range)_currentWorksheet.Cells[_currentRow, _currentColumn];

            Excel.Range documentDimensions = _currentWorksheet.Range[documentStart, documentEnd];

            Excel.ChartObjects chartObjects = (Excel.ChartObjects)_currentWorksheet.ChartObjects();

            Excel.ChartObject chartObject = chartObjects.Add((double)documentDimensions.Width, (double)documentDimensions.Height, widthPX, heightPX);

            Excel.Chart chart = chartObject.Chart;

            Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();

            Excel.Series series = seriesCollection.NewSeries();

            series.Values = yvalues;

            series.XValues = xvalues;

            series.Name = seriesName;

            Excel.Range counter = (Excel.Range)_currentWorksheet.Cells[++_currentRow, _currentColumn];
            
            while (heightPX > (double)(_currentWorksheet.Range[documentEnd, counter]).Height)
            {
                counter = (Excel.Range)_currentWorksheet.Cells[++_currentRow, _currentColumn];
            }

            Deallocate(documentStart, documentEnd, documentDimensions, chartObjects, chartObject, chart, seriesCollection, series);

            return this;
        }

        public ExcelWriter AppendText(string text, int columnOffset = 0, int rowOffset = 0)
        {
            columnOffset = columnOffset > 0 ? columnOffset : 0;
            rowOffset = rowOffset > 0 ? rowOffset : 0;
            Excel.Range cell = (Excel.Range)_currentWorksheet.Cells[rowOffset + _currentRow, columnOffset + _currentColumn++];
            cell.Value = text;
            CellStyle.ApplyStyle(cell);
            return this;
        }

        public ExcelWriter SetCurrentSheetName(string name)
        {
            _currentWorksheet.Name = name;

            return this;
        }

        public ExcelWriter NewSheet(string name = "")
        {
            _currentWorksheet = (Excel.Worksheet)_worksheets.Add(NoValue, _currentWorksheet, NoValue, NoValue);
            if (name != string.Empty) _currentWorksheet.Name = name;
            _currentColumn = BaseColumn;
            _currentRow = BaseRow;
            return this;
        }

        #endregion

        #region Inherited Methods

        public override void Write(string targetPath)
        {
            _workbook.SaveAs(
                targetPath,
                Excel.XlFileFormat.xlWorkbookNormal,
                NoValue,
                NoValue,
                NoValue,
                NoValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                NoValue,
                NoValue,
                NoValue,
                NoValue,
                NoValue);
            _workbook.Close(false, NoValue, NoValue);
            _application.Quit();

        }

        public override void Dispose()
        {
            Deallocate(_application, _workbooks, _workbook, _worksheets, _currentWorksheet);
        }

        #endregion

        private class FieldData
        {
            public PropertyInfo Property { get; }
            public string FieldName { get; }
            public string Format { get; }

            public FieldData(PropertyInfo property) : this(property, property.Name) { }

            public FieldData(PropertyInfo property, string fieldName)
            {
                Property = property;
                FieldName = fieldName;
            }

            public FieldData(bool include)
            {
            }
        }

        public abstract class WriterStyle
        {
            public abstract void ApplyStyle(Excel.Range range);
        }

        public class NoStyle : WriterStyle
        {
            #region Singleton

            private static NoStyle instance;

            private static readonly object objLock = new object();

            public static NoStyle Instance
            {
                get
                {
                    if (instance == null)
                    {
                        lock (objLock)
                        {
                            if (instance == null)
                                instance = new NoStyle();
                        }
                    }

                    return instance;
                }
            }

            #endregion

            private NoStyle() { }

            public override void ApplyStyle(Excel.Range range) { }
        }

        public class BasicWriterStyle : WriterStyle
        {
            public string FontName { get; }

            public int FontSize { get; }

            public Color FontColor { get; }

            public Color FillColor { get; }

            public bool Bold { get; }

            public bool Italic { get; }

            public bool Underlined { get; }

            public bool StrikeThrough { get; }

            public int CellHeight { get; }

            public int CellWidth { get; }

            public HorizontalAlignment HorizontalAlignment { get; }

            public VerticalAlignment VerticalAlignment { get; }

            public BasicWriterStyle(
                string fontName = "Arial",
                int fontSize = 12,
                bool bold = false,
                bool italic = false,
                bool underlined = false,
                bool strikeThrough = false,
                int cellWidth = 0,
                int cellHeight = 0,
                HorizontalAlignment halign = HorizontalAlignment.Variable,
                VerticalAlignment valign = VerticalAlignment.Bottom)
                : this(fontName, fontSize, bold, italic, underlined, strikeThrough, cellWidth, cellHeight, halign, valign, Color.Black, Color.Transparent) { }

            public BasicWriterStyle(
                string fontName,
                int fontSize,
                bool bold,
                bool italic,
                bool underlined,
                bool strikeThrough,
                int cellWidth, 
                int cellHeight, 
                HorizontalAlignment halign,
                VerticalAlignment valign,
                Color fontColor,
                Color fillColor)
            {
                FontName = fontName;
                FontSize = fontSize;
                FontColor = fontColor;
                FillColor = fillColor;
                Bold = bold;
                Italic = italic;
                Underlined = underlined;
                StrikeThrough = strikeThrough;
                CellWidth = cellWidth;
                CellHeight = cellHeight;
                HorizontalAlignment = halign;
                VerticalAlignment = valign;
            }

            public override void ApplyStyle(Excel.Range range)
            {
                if (FontName != string.Empty) range.Font.Name = FontName;
                if (FontSize > 0) range.Font.Size = FontSize;
                range.Font.Color = ColorTranslator.ToOle(FontColor);
                if (FillColor != Color.Transparent) range.Interior.Color = ColorTranslator.ToOle(FillColor);
                range.Font.Bold = Bold;
                range.Font.Italic = Italic;
                range.Font.Underline = Underlined;
                range.Font.Strikethrough = StrikeThrough;
                if (CellWidth > 0) range.ColumnWidth = CellWidth;
                if (CellHeight > 0) range.EntireRow.RowHeight = CellHeight;
                range.HorizontalAlignment = horizontalAlignmentMaps[HorizontalAlignment];
                range.VerticalAlignment = verticalAlignmentMaps[VerticalAlignment];
                

            }
        }

        public class PredefinedWriterStyle : WriterStyle
        {

            private readonly string styleName;

            public PredefinedWriterStyle(string styleName) => this.styleName = styleName;

            public override void ApplyStyle(Excel.Range range)
            {
                range.Style = styleName;
            }
        }

        public class TableStyle
        {
            public WriterStyle TitleStyle { get; set; }
            public WriterStyle FieldNameStyle { get; set; }
            public WriterStyle BodyStyle { get; set; }

            public TableStyle()
            {
                BodyStyle = FieldNameStyle = TitleStyle = NoStyle.Instance;
            }
        }

        public class StripedRowWriterStyle : WriterStyle
        {
            private readonly Color[] stripeColors;

            public StripedRowWriterStyle(Color[] stripeColors) => this.stripeColors = stripeColors;

            public override void ApplyStyle(Excel.Range range)
            {
                Color[] fillColorCache = stripeColors;

                int stripeCount = fillColorCache.Length - 1;

                int rowCount = range.Rows.Count;

                int relativeIndex = 0;

                for (int currentRow = 1; currentRow <= rowCount; ++currentRow)
                {
                    Excel.Range current = (Excel.Range)range.Rows[currentRow];

                    current.Interior.Color = ColorTranslator.ToOle(fillColorCache[relativeIndex]);

                    relativeIndex = (relativeIndex < stripeCount) ? ++relativeIndex : 0;
                }
            }
        }
    }

    public enum HorizontalAlignment
    {
        Variable,
        Left,
        Right,
        Middle,
        SpacedIn,
        SpacedOut
    }

    public enum VerticalAlignment
    {
        Top,
        Bottom,
        Center,
        SpacedIn,
        SpacedOut
    }
}
