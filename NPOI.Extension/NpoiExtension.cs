// Copyright (c) RigoFunc (xuyingting). All rights reserved.

namespace NPOI.Extension {
    using System;
    using System.Collections.Generic;
    using System.IO;

    // NPOI
    using HPSF;
    using HSSF.UserModel;
    using HSSF.Util;
    using SS.UserModel;
    using SS.Util;

    /// <summary>
    /// Defines some extensions of NPOI.
    /// </summary>
    public static class NpoiExtension {
        public static byte[] ToExcelContent<T>(this IEnumerable<T> source) {
            var book = source.ToWorkbook();

            using (var ms = new MemoryStream()) {
                book.Write(ms);
                return ms.ToArray();
            }
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string fileName) {
            var book = source.ToWorkbook();

            // Write the stream data of workbook to file
            using (var stream = new FileStream(fileName, FileMode.OpenOrCreate)) {
                book.Write(stream);
            }
        }

        internal static IWorkbook ToWorkbook<T>(this IEnumerable<T> source) {
            if (source == null) {
                throw new ArgumentNullException(nameof(source));
            }

            var properties = typeof(T).GetProperties();

            // find out the attribute
            var attributes = new ColumnAttribute[properties.Length];
            for (var j = 0; j < properties.Length; j++) {
                var property = properties[j];
                var attrs = property.GetCustomAttributes(typeof(ColumnAttribute), true) as ColumnAttribute[];
                if (attrs != null && attrs.Length > 0) {
                    attributes[j] = attrs[0];
                }
                else {
                    attributes[j] = null;
                }
            }

            // Init work book.
            var workbook = InitializeWorkbook();

            // new sheet.
            var sheet = workbook.CreateSheet();

            // column (first row) title style
            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = HSSFColor.White.Index;
            style.FillPattern = FillPattern.Bricks;
            style.FillBackgroundColor = HSSFColor.Grey40Percent.Index;

            var rowIndex = 1;
            foreach (var item in source) {
                var row = sheet.CreateRow(rowIndex);
                for (var i = 0; i < properties.Length; i++) {
                    var property = properties[i];
                    var excel = attributes[i];
                    if (excel == null)
                        continue;

                    var value = property.GetValue(item);
                    if (value != null) {
                        var cell = row.CreateCell(excel.Index);
                        if (value is ValueType) {
                            if (property.PropertyType == typeof(bool)) {
                                cell.SetCellValue((bool)value);
                            }
                            else if (property.PropertyType == typeof(DateTime)) {
                                cell.SetCellValue(Convert.ToDateTime(value).ToString("yyyy-MM-dd"));
                            }
                            else {
                                cell.SetCellValue(Convert.ToDouble(value));
                            }
                        }
                        else {
                            cell.SetCellValue(value + "");
                        }
                    }
                    else {
                        row.CreateCell(excel.Index).SetCellValue(string.Empty);
                    }
                }
                rowIndex++;
            }

            // first row (column title)
            var row1 = sheet.CreateRow(0);
            for (var i = 0; i < properties.Length; i++) {
                var property = properties[i];
                var excel = attributes[i];
                if (excel == null)
                    continue;

                var cell = row1.CreateCell(excel.Index);
                cell.CellStyle = style;

                cell.SetCellValue(excel.Title);
            }

            // total row
            if (rowIndex > 0) {
                var totalRow = sheet.CreateRow(rowIndex);
                var cell = totalRow.CreateCell(0);
                cell.CellStyle = style;
                cell.SetCellValue("Total");
                foreach (var item in attributes) {
                    if (item.AllowSum) {
                        cell = totalRow.CreateCell(item.Index);
                        cell.CellFormula = $"SUM({GetCellPosition(1, item.Index)}:{GetCellPosition(rowIndex - 1, item.Index)})";
                    }
                }
            }

            // merge cell style
            var style2 = workbook.CreateCellStyle();
            style2.VerticalAlignment = VerticalAlignment.Center;

            // merge cells
            for (var j = 0; j < attributes.Length; j++) {
                var excel = attributes[j];
                var previous = "";
                int rowspan = 0, row = 1;
                if (excel.AllowMerge) {
                    for (row = 1; row < rowIndex; row++) {
                        var value = sheet.GetRow(row).Cells[excel.Index].StringCellValue;
                        if (previous == value && !string.IsNullOrEmpty(value)) {
                            rowspan++;
                        }
                        else {
                            if (rowspan > 1) {
                                sheet.GetRow(row - rowspan).Cells[excel.Index].CellStyle = style2;
                                sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, excel.Index, excel.Index));
                            }
                            rowspan = 1;
                            previous = value;
                        }
                    }

                    // in what case? -> all rows need to be merged
                    if (rowspan > 1) {
                        sheet.GetRow(row - rowspan).Cells[excel.Index].CellStyle = style2;
                        sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, excel.Index, excel.Index));
                    }
                }
            }

            // set the freeze
            var fattrs = typeof(T).GetCustomAttributes(typeof(FreezeAttribute), true) as FreezeAttribute[];
            if (fattrs != null && fattrs.Length > 0) {
                var freeze = fattrs[0];
                sheet.CreateFreezePane(freeze.ColSplit, freeze.RowSplit, freeze.LeftMostColumn, freeze.TopRow);
            }

            // set the auto filter
            var filters = typeof(T).GetCustomAttributes(typeof(FilterAttribute), true) as FilterAttribute[];
            if (filters != null && filters.Length > 0) {
                var filter = filters[0];
                sheet.SetAutoFilter(new CellRangeAddress(filter.FirstRow, filter.LastRow ?? rowIndex, filter.FirstCol, filter.LastCol));
            }

            // autosize the all columns
            for (int i = 0; i < properties.Length; i++) {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        private static HSSFWorkbook InitializeWorkbook() {
            var hssfworkbook = new HSSFWorkbook();

            //Create a entry of DocumentSummaryInformation
            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "RigoFunc (xyting)";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //Create a entry of SummaryInformation
            var si = PropertySetFactory.CreateSummaryInformation();
            si.Author = "RigoFunc (xyting)";
            si.Subject = "NPOI Extension";
            hssfworkbook.SummaryInformation = si;

            return hssfworkbook;
        }

        private static string GetCellPosition(int row, int col) {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }
    }
}
