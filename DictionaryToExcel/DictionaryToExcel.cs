using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ougha.Entities;

namespace Ougha
{
    public class DictionaryToExcel
    {
        public static byte[] CreateExcel(List<Dictionary<string, string>> items)
        {
        
            return _createExcel(items);
        }
        public static byte[] CreateExcel<T>(T items) where T : List<TabSheet>
        {

            return _createExcel(items);
        }


        [Obsolete]
        static byte[] _createExcel(List<Dictionary<string, string>> items)
        {
            using (var ms = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet();
                    sheet.Id = document.WorkbookPart.GetIdOfPart(worksheetPart);
                    sheet.SheetId = 1;
                    sheet.Name = $"Sheet1";
                    sheets.Append(sheet);

                    _addStyles(document);

                    if (items.Any())
                    {
                        var cols = worksheetPart.Worksheet.AppendChild(new Columns());
                        cols.Append(items.Select(x => new Column() { Min = (uint)(items.IndexOf(x) + 1), Max = (uint)(items.IndexOf(x) + 1), Width = 20, CustomWidth = true, BestFit = true }));

                        var rows = _getRows(items);
                        var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                        sheetData.Append(rows);

                        Dictionary<string, string> first = null;
                        foreach (var item in items)
                        {
                            first = item;
                            break;
                        }

                        worksheetPart.Worksheet.Append(new AutoFilter() { Reference = $"A1:{_getColReference(first.Count - 1)}{rows.Length}" });
                    }

                    workbookpart.Workbook.Save();
                }
                return ms.ToArray();
            }
        }
        static byte[] _createExcel<T>(T items) where T : List<TabSheet>
        {
            using (var ms = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();
                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    _addSheets(sheets, workbookpart, document,items);
                    workbookpart.Workbook.Save();
                }
                return ms.ToArray();
            }
        }
        static void _addSheets(Sheets sheets, WorkbookPart workbookPart, SpreadsheetDocument document, List<TabSheet> tabSheets = null)
        {
            if (tabSheets != null)
            {
                foreach (var tab in tabSheets)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    var sheet = new Sheet();
                    sheet.Id = document.WorkbookPart.GetIdOfPart(worksheetPart);
                    UInt32 sheetId = (UInt32)tabSheets.IndexOf(tab) + 2;
                    sheet.SheetId = sheetId;
                    sheet.Name = tab.Name;
                    sheets.Append(sheet);

                    //_addStyles(document);

                    if (tab.Properties.Any())
                    {
                        var cols = worksheetPart.Worksheet.AppendChild(new Columns());
                        cols.Append(tab.Properties.Select(x => new Column() { Min = (uint)(tab.Properties.IndexOf(x) + 1), Max = (uint)(tab.Properties.IndexOf(x) + 1), Width = 20, CustomWidth = true, BestFit = true }));

                        var rows = _getRows(tab.Properties);
                        var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                        sheetData.Append(rows);

                        Dictionary<string, string> first = null;
                        foreach (var item in tab.Properties)
                        {
                            first = item;
                            break;
                        }

                        worksheetPart.Worksheet.Append(new AutoFilter() { Reference = $"A1:{_getColReference(first.Count - 1)}{rows.Length}" });
                    }
                }
            }
        }
        static void _addStyles(SpreadsheetDocument document)
        {
            var stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            // fonts
            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            var font1 = stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            font1.Append(new Bold());
            font1.Append(new Color() { Rgb = HexBinaryValue.FromString("FFFFFFFF") });
            stylesPart.Stylesheet.Fonts.Count = (uint)stylesPart.Stylesheet.Fonts.ChildElements.Count;

            // fills
            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            var fill2 = stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill() { PatternType = PatternValues.Solid } });
            fill2.PatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF4F81BD") };
            fill2.PatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            stylesPart.Stylesheet.Fills.Count = (uint)stylesPart.Stylesheet.Fills.ChildElements.Count;

            // borders
            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.AppendChild(new Border());
            stylesPart.Stylesheet.Borders.Count = (uint)stylesPart.Stylesheet.Borders.ChildElements.Count;

            // NumberingFormats
            //uint iExcelIndex = 164;
            stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.Count = (uint)stylesPart.Stylesheet.NumberingFormats.ChildElements.Count;

            // cell style formats
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.Count = 1;

            // cell styles
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            // header style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center });
            // datetime style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true });
            stylesPart.Stylesheet.CellFormats.Count = (uint)stylesPart.Stylesheet.CellFormats.ChildElements.Count;

            stylesPart.Stylesheet.Save();
        }

        static Row[] _getRows<T>(T items) where T : List<Dictionary<string, string>>
        {


            var rows = new List<Row>();

            if (items == null && !items.Any())
            {
                return rows.ToArray();
            }

            var headers = new List<Cell>();
            var j = 1;
            foreach (var item in items.FirstOrDefault())
            {
                headers.Add(new Cell
                {
                    CellReference = _getColReference(j),
                    CellValue = new CellValue(item.Key),
                    DataType = CellValues.String,
                    StyleIndex = 0u,
                });
                j++;
            }

            var headerCells = headers.ToArray();

            var headerRow = new Row() { RowIndex = 1 };
            headerRow.Append(headerCells);
            rows.Add(headerRow);

            var i = 2;

            foreach (var item in items)
            {
                var cells = new List<Cell>();
                var row = new Row() { RowIndex = (uint)i++ };
                j = 1;
                foreach (var valueRow in item)
                {
                    cells.Add(new Cell
                    {
                        CellReference = _getColReference(j),
                        CellValue = new CellValue(valueRow.Value),
                        DataType = CellValues.String,
                        StyleIndex = 1
                    });
                    j++;
                }


                row.Append(cells);
                rows.Add(row);


            }
            return rows.ToArray();
        }

        static Cell _getCell(string reference, object value)
        {
            var dataType = _getCellType(value);
            return new Cell
            {
                CellReference = reference,
                CellValue = _getCellValue(value),
                DataType = dataType,
                StyleIndex = dataType == CellValues.Date ? 2 : 0u,
            };
        }

        static CellValue _getCellValue(object value)
        {
            if (value == null) return new CellValue();

            var type = value.GetType();

            if (type == typeof(bool))
                return new CellValue((bool)value ? "1" : "0");

            if (type == typeof(DateTime))
                return new CellValue(((DateTime)value).ToString("s", _cultureInfo));

            if (type == typeof(DateTimeOffset))
                return new CellValue(((DateTimeOffset)value).ToString("s", _cultureInfo));

            if (type == typeof(double))
                return new CellValue(((double)value).ToString(_cultureInfo));

            if (type == typeof(decimal))
                return new CellValue(((decimal)value).ToString(_cultureInfo));

            if (type == typeof(float))
                return new CellValue(((float)value).ToString(_cultureInfo));

            return new CellValue(value.ToString());
        }

        static CellValues _getCellType(object value)
        {
            var type = value?.GetType();

            if (type == typeof(bool))
                return CellValues.Boolean;

            if (_numericTypes.Contains(type))
                return CellValues.Number;

            if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
                return CellValues.Date;

            return CellValues.String;
        }

        static string _getColReference(int index)
        {
            var result = new List<char>();
            while (index >= _digits.Length)
            {
                int remainder = index % _digits.Length;
                index = index / _digits.Length - 1;
                result.Add(_digits[remainder]);
            }
            result.Add(_digits[index]);
            result.Reverse();
            return new string(result.ToArray());
        }

        static HashSet<Type> _numericTypes = new HashSet<Type>
        {
            typeof(short),
            typeof(ushort),
            typeof(int),
            typeof(uint),
            typeof(long),
            typeof(ulong),
            typeof(double),
            typeof(decimal),
            typeof(float),
        };

        static CultureInfo _cultureInfo = CultureInfo.GetCultureInfo("en-US");

        static string _digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    }
}
