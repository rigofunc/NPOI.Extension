// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace NPOI.Extension
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    // NPOI
    using HPSF;
    using HSSF.UserModel;
    using HSSF.Util;
    using SS.UserModel;
    using SS.Util;
    using XSSF.UserModel;

    /// <summary>
    /// Defines some extensions for <see cref="IEnumerable{T}"/> that using NPOI to provides excel functionality.
    /// </summary>
    public static class IEnumerableNpoiExtensions
    {
        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, string sheetName = "sheet0")
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            var book = source.ToWorkbook(null, sheetName);

            using (var ms = new MemoryStream())
            {
                book.Write(ms);
                return ms.ToArray();
            }
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string excelFile, string sheetName = "sheet0") where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            if (Path.GetExtension(excelFile).Equals(".xls"))
            {
                Excel.Setting.UserXlsx = false;
            }
            else if (Path.GetExtension(excelFile).Equals(".xlsx"))
            {
                Excel.Setting.UserXlsx = true;
            }
            else
            {
                throw new NotSupportedException($"not an excel file extension (*.xls | *.xlsx) {excelFile}");
            }

            var book = source.ToWorkbook(excelFile, sheetName);

            // Write the stream data of workbook to file
            using (var stream = new FileStream(excelFile, FileMode.OpenOrCreate, FileAccess.Write))
            {
                book.Write(stream);
            }
        }

        internal static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, string excelFile, string sheetName)
        {
            // can static properties or only instance properties?
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            // find out the attributes
            var haventCols = true;
            var attributes = new ColumnAttribute[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];
                var attrs = property.GetCustomAttributes(typeof(ColumnAttribute), true) as ColumnAttribute[];
                if (attrs != null && attrs.Length > 0)
                {
                    attributes[j] = attrs[0];

                    // attribute configure first(Hight Priority)
                    haventCols = false;
                }
                else
                {
                    attributes[j] = null;
                }
            }

            // init work book.
            var workbook = InitializeWorkbook(excelFile);

            // new sheet
            var sheet = workbook.CreateSheet(sheetName);

            // cache for datetime format
            ICellStyle dateCellStyle = null;

            var rowIndex = 1;
            foreach (var item in source)
            {
                var row = sheet.CreateRow(rowIndex);
                for (var i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];

                    int index = i;
                    if (!haventCols)
                    {
                        var column = attributes[i];
                        if (column == null)
                            continue;
                        else
                            index = column.Index;
                    }

                    var value = property.GetValue(item, null);
                    var cell = row.CreateCell(index);
                    if (value is ValueType)
                    {
                        if (value == null)
                        {
                            // do nothing here?
                            continue;
                        }

                        if (property.PropertyType.UnwrapNullableType() == typeof(bool))
                        {
                            cell.SetCellValue((bool)value);
                        }
                        else if (property.PropertyType.UnwrapNullableType() == typeof(DateTime))
                        {
                            if (dateCellStyle == null)
                            {
                                // create the cache.
                                dateCellStyle = workbook.CreateCellStyle();

                                var dateFormat = workbook.CreateDataFormat();

                                dateCellStyle.DataFormat = dateFormat.GetFormat(Excel.Setting.DateFormatter);
                            }

                            cell.CellStyle = dateCellStyle;

                            cell.SetCellValue(Convert.ToDateTime(value));
                        }
                        else if (property.PropertyType.UnwrapNullableType() == typeof(Guid))
                        {
                            cell.SetCellValue(Convert.ToString(value));
                        }
                        else
                        {
                            cell.SetCellValue(Convert.ToDouble(value));
                        }
                    }
                    else
                    {
                        // even if: null + ""
                        cell.SetCellValue(value + "");
                    }
                }
                rowIndex++;
            }

            if (!haventCols)
            {
                // merge cell style
                var vStyle = workbook.CreateCellStyle();
                vStyle.VerticalAlignment = VerticalAlignment.Center;

                // merge cells
                for (var j = 0; j < attributes.Length; j++)
                {
                    var column = attributes[j];
                    if (column == null)
                    {
                        continue;
                    }

                    var previous = "";
                    int rowspan = 0, row = 1;
                    if (column.AllowMerge)
                    {
                        for (row = 1; row < rowIndex; row++)
                        {
                            var value = sheet.GetRow(row).Cells[column.Index].StringCellValue;
                            if (previous == value && !string.IsNullOrEmpty(value))
                            {
                                rowspan++;
                            }
                            else
                            {
                                if (rowspan > 1)
                                {
                                    sheet.GetRow(row - rowspan).Cells[column.Index].CellStyle = vStyle;
                                    sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, column.Index, column.Index));
                                }
                                rowspan = 1;
                                previous = value;
                            }
                        }

                        // in what case? -> all rows need to be merged
                        if (rowspan > 1)
                        {
                            sheet.GetRow(row - rowspan).Cells[column.Index].CellStyle = vStyle;
                            sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, column.Index, column.Index));
                        }
                    }
                }
            }

            // column (first row) title style
            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = HSSFColor.White.Index;
            style.FillPattern = FillPattern.Bricks;
            style.FillBackgroundColor = HSSFColor.Grey40Percent.Index;

            // first row (column title)
            var row1 = sheet.CreateRow(0);
            for (var i = 0; i < properties.Length; i++)
            {
                var property = properties[i];

                var title = property.Name;
                int index = i;
                if (!haventCols)
                {
                    var column = attributes[i];
                    if (column == null)
                        continue;
                    else
                    {
                        index = column.Index;
                        // if not title, using property name as title.
                        if (!string.IsNullOrEmpty(column.Title))
                        {
                            title = column.Title;
                        }
                    }
                }

                var cell = row1.CreateCell(index);
                cell.CellStyle = style;
                cell.SetCellValue(title);
            }

            if (rowIndex > 0)
            {
                // statistics row
                var statistics = typeof(T).GetCustomAttributes(typeof(StatisticsAttribute), true) as StatisticsAttribute[];
                if (statistics != null && statistics.Length > 0)
                {
                    var first = statistics[0];
                    var lastRow = sheet.CreateRow(rowIndex);
                    var cell = lastRow.CreateCell(0);
                    cell.SetCellValue(first.Name);
                    foreach (var column in first.Columns)
                    {
                        cell = lastRow.CreateCell(column);
                        cell.CellFormula = $"{first.Formula}({GetCellPosition(1, column)}:{GetCellPosition(rowIndex - 1, column)})";
                    }
                }

                // set the freeze
                var fattrs = typeof(T).GetCustomAttributes(typeof(FreezeAttribute), true) as FreezeAttribute[];
                if (fattrs != null && fattrs.Length > 0)
                {
                    var freeze = fattrs[0];
                    sheet.CreateFreezePane(freeze.ColSplit, freeze.RowSplit, freeze.LeftMostColumn, freeze.TopRow);
                }

                // set the auto filter
                var filters = typeof(T).GetCustomAttributes(typeof(FilterAttribute), true) as FilterAttribute[];
                if (filters != null && filters.Length > 0)
                {
                    var filter = filters[0];
                    sheet.SetAutoFilter(new CellRangeAddress(filter.FirstRow, filter.LastRow ?? rowIndex, filter.FirstCol, filter.LastCol));
                }
            }

            // autosize the all columns
            for (int i = 0; i < properties.Length; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        private static IWorkbook InitializeWorkbook(string excelFile)
        {
            var setting = Excel.Setting;
            if (setting.UserXlsx)
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        return new XSSFWorkbook(file);
                    }
                }
                else
                {
                    return new XSSFWorkbook();
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        return new HSSFWorkbook(file);
                    }
                }
                else
                {
                    var hssf = new HSSFWorkbook();

                    var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                    dsi.Company = setting.Company;
                    hssf.DocumentSummaryInformation = dsi;

                    var si = PropertySetFactory.CreateSummaryInformation();
                    si.Author = setting.Author;
                    si.Subject = setting.Subject;
                    hssf.SummaryInformation = si;

                    return hssf;
                }
            }
        }

        private static string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }
    }
}
