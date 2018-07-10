// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;
    using NPOI.HPSF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Defines some extensions for <see cref="IEnumerable{T}"/> that using NPOI to provides excel functionality.
    /// </summary>
    public static class IEnumerableNpoiExtensions
    {
        private static IFormulaEvaluator _formulaEvaluator;

        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            return ToExcel(source, null, s => sheetName, maxRowsPerSheet, overwrite);
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string excelFile, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            //TODO check the file's path is valid
            ToExcel(source, excelFile, s => sheetName, maxRowsPerSheet, overwrite);
        }

        public static byte[] ToExcel<T>(this IEnumerable<T> source, string excelFile, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
           where T : class
        {
            return ToExcel(source, excelFile, Excel.Setting, sheetSelector, maxRowsPerSheet, overwrite);
        }

        public static byte[] ToExcel<T>(this IEnumerable<T> source, string excelFile, ExcelSetting excelSetting, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            bool isVolatile = string.IsNullOrWhiteSpace(excelFile);
            if (!isVolatile)
            {
                var extension = Path.GetExtension(excelFile);
                if (extension.Equals(".xls"))
                {
                    excelSetting.UseXlsx = false;
                }
                else if (extension.Equals(".xlsx"))
                {
                    excelSetting.UseXlsx = true;
                }
                else
                {
                    throw new NotSupportedException($"not an excel file (*.xls | *.xlsx) extension: {extension}");
                }
            }
            else
            {
                excelFile = null;
            }

            IWorkbook book = InitializeWorkbook(excelFile, excelSetting);
            using (Stream ms = isVolatile ? (Stream)new MemoryStream() : new FileStream(excelFile, FileMode.OpenOrCreate, FileAccess.Write))
            {
                IEnumerable<byte> output = Enumerable.Empty<byte>();
                foreach (var sheet in source.AsQueryable().GroupBy(sheetSelector))
                {
                    int sheetIndex = 0;
                    var content = sheet.Select(row => row);
                    while (content.Any())
                    {
                        book = content.Take(maxRowsPerSheet).ToWorkbook(book, sheet.Key + (sheetIndex > 0 ? "_" + sheetIndex.ToString() : ""), overwrite);
                        sheetIndex++;
                        content = content.Skip(maxRowsPerSheet);
                    }
                }
                book.Write(ms);
                return isVolatile ? ((MemoryStream)ms).ToArray() : null;
            }
        }

        public static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, string sheetName = "sheet0") where T : class
            => ToWorkbook<T>(source, Excel.Setting, sheetName);

        public static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, ExcelSetting excelSetting, string sheetName = "sheet0") where T : class
            => ToWorkbook<T>(source, InitializeWorkbook(null, excelSetting), excelSetting, sheetName, false);

        public static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, IWorkbook workbook, string sheetName = "sheet0", bool overwrite = false) where T : class
            => ToWorkbook<T>(source, workbook, Excel.Setting, sheetName, overwrite);

        public static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, IWorkbook workbook, ExcelSetting excelSetting, string sheetName = "sheet0", bool overwrite = false)
            where T : class
        {
            if (null == source) throw new ArgumentNullException(nameof(source));
            if (null == workbook) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException($"sheet name cannot be null or whitespace", nameof(sheetName));

            // TODO: can static properties or only instance properties?
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            bool fluentConfigEnabled = false;
            // get the fluent config
            if (excelSetting.FluentConfigs.TryGetValue(typeof(T), out var fluentConfig))
            {
                fluentConfigEnabled = true;

                // adjust the auto index.
                (fluentConfig as FluentConfiguration<T>)?.AdjustAutoIndex();
            }

            // find out the configurations
            var propertyConfigurations = new PropertyConfiguration[properties.Length];
            for (var i = 0; i < properties.Length; i++)
            {
                var property = properties[i];

                // get the property config
                if (fluentConfigEnabled && fluentConfig.PropertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    propertyConfigurations[i] = pc;
                }
                else
                {
                    propertyConfigurations[i] = null;
                }
            }

            // TODO check the sheet's name is valid
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            else
            {
                // doesn't override the exist sheet if not required
                if (!overwrite) sheet = workbook.CreateSheet();
            }

            // cache cell styles
            var cellStyles = new Dictionary<int, ICellStyle>();

            // title row cell style
            ICellStyle titleStyle = null;
            if (excelSetting.TitleCellStyleApplier != null)
            {
                titleStyle = workbook.CreateCellStyle();
                var font = workbook.CreateFont();
                excelSetting.TitleCellStyleApplier(titleStyle, font);
            }

            var titleRow = sheet.CreateRow(0);
            var rowIndex = 1;
            foreach (var item in source)
            {
                var row = sheet.CreateRow(rowIndex);
                for (var i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];

                    int index = i;
                    var config = propertyConfigurations[i];
                    if (config != null)
                    {
                        if (config.IsExportIgnored)
                            continue;

                        index = config.Index;

                        if (index < 0)
                            throw new Exception($"The excel cell index value cannot be less then '0' for the property: {property.Name}, see HasExcelIndex(int index) methods for more informations.");
                    }

                    // this is the first time.
                    if (rowIndex == 1)
                    {
                        // if not title, using property name as title.
                        var title = property.Name;
                        if (!string.IsNullOrEmpty(config?.Title))
                        {
                            title = config.Title;
                        }

                        if (!string.IsNullOrEmpty(config?.Formatter))
                        {
                            try
                            {
                                var style = workbook.CreateCellStyle();

                                var dataFormat = workbook.CreateDataFormat();

                                style.DataFormat = dataFormat.GetFormat(config.Formatter);

                                cellStyles[i] = style;
                            }
                            catch (Exception ex)
                            {
                                // the formatter isn't excel supported formatter
                                System.Diagnostics.Debug.WriteLine(ex.ToString());
                            }
                        }

                        var titleCell = titleRow.CreateCell(index);
                        if (titleStyle != null)
                        {
                            titleCell.CellStyle = titleStyle;
                        }
                        titleCell.SetCellValue(title);
                    }

                    var unwrapType = property.PropertyType.UnwrapNullableType();

                    var value = property.GetValue(item, null);

                    // give a chance to the value converter even though value is null.
                    if (config?.CellValueConverter != null)
                    {
                        value = config.CellValueConverter(rowIndex, index, value);
                        if (value == null)
                            continue;

                        unwrapType = value.GetType().UnwrapNullableType();
                    }

                    if (value == null)
                        continue;

                    var cell = row.CreateCell(index);
                    if (cellStyles.TryGetValue(i, out var cellStyle))
                    {
                        cell.CellStyle = cellStyle;
                    }
                    else if (!string.IsNullOrEmpty(config?.Formatter) && value is IFormattable fv)
                    {
                        // the formatter isn't excel supported formatter, but it's a C# formatter.
                        // The result is the Excel cell data type become String.
                        cell.SetCellValue(fv.ToString(config.Formatter, CultureInfo.CurrentCulture));

                        continue;
                    }

                    if (unwrapType == typeof(bool))
                    {
                        cell.SetCellValue((bool)value);
                    }
                    else if (unwrapType == typeof(DateTime))
                    {
                        cell.SetCellValue(Convert.ToDateTime(value));
                    }
                    else if (unwrapType.IsInteger() ||
                             unwrapType == typeof(decimal) ||
                             unwrapType == typeof(double) ||
                             unwrapType == typeof(float))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else
                    {
                        cell.SetCellValue(value.ToString());
                    }
                }

                rowIndex++;
            }

            // merge cells
            var mergableConfigs = propertyConfigurations.Where(c => c != null && c.AllowMerge).ToList();
            if (mergableConfigs.Any())
            {
                // merge cell style
                var vStyle = workbook.CreateCellStyle();
                vStyle.VerticalAlignment = VerticalAlignment.Center;

                foreach (var config in mergableConfigs)
                {
                    object previous = null;
                    int rowspan = 0, row = 1;
                    for (row = 1; row < rowIndex; row++)
                    {
                        var value = sheet.GetRow(row).GetCellValue(config.Index, _formulaEvaluator);
                        if (object.Equals(previous, value) && value != null)
                        {
                            rowspan++;
                        }
                        else
                        {
                            if (rowspan > 1)
                            {
                                sheet.GetRow(row - rowspan).Cells[config.Index].CellStyle = vStyle;
                                sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, config.Index, config.Index));
                            }
                            rowspan = 1;
                            previous = value;
                        }
                    }

                    // in what case? -> all rows need to be merged
                    if (rowspan > 1)
                    {
                        sheet.GetRow(row - rowspan).Cells[config.Index].CellStyle = vStyle;
                        sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, config.Index, config.Index));
                    }
                }
            }

            if (rowIndex > 1 && fluentConfigEnabled)
            {
                var statistics = fluentConfig.StatisticsConfigurations;
                var filterConfigs = fluentConfig.FilterConfigurations;
                var freezeConfigs = fluentConfig.FreezeConfigurations;

                // statistics row
                foreach (var item in statistics)
                {
                    var lastRow = sheet.CreateRow(rowIndex);
                    var cell = lastRow.CreateCell(0);
                    cell.SetCellValue(item.Name);
                    foreach (var column in item.Columns)
                    {
                        cell = lastRow.CreateCell(column);

                        // set the same cell style
                        cell.CellStyle = sheet.GetRow(rowIndex - 1)?.GetCell(column)?.CellStyle;

                        // set the cell formula
                        cell.CellFormula = $"{item.Formula}({GetCellPosition(1, column)}:{GetCellPosition(rowIndex - 1, column)})";
                    }

                    rowIndex++;
                }

                // set the freeze
                foreach (var freeze in freezeConfigs)
                {
                    sheet.CreateFreezePane(freeze.ColSplit, freeze.RowSplit, freeze.LeftMostColumn, freeze.TopRow);
                }

                // set the auto filter
                foreach (var filter in filterConfigs)
                {
                    sheet.SetAutoFilter(new CellRangeAddress(filter.FirstRow, filter.LastRow ?? rowIndex, filter.FirstCol, filter.LastCol));
                }
            }

            // autosize the all columns
            if (excelSetting.AutoSizeColumnsEnabled)
            {
                for (int i = 0; i < properties.Length; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }

            return workbook;
        }

        private static IWorkbook InitializeWorkbook(string excelFile, ExcelSetting excelSetting = null)
        {
            var setting = excelSetting ?? Excel.Setting;
            if (setting.UseXlsx)
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        var workbook = new XSSFWorkbook(file);

                        _formulaEvaluator = new XSSFFormulaEvaluator(workbook);

                        return workbook;
                    }
                }
                else
                {
                    var workbook = new XSSFWorkbook();

                    _formulaEvaluator = new XSSFFormulaEvaluator(workbook);

                    var props = workbook.GetProperties();
                    props.CoreProperties.Creator = setting.Author;
                    props.CoreProperties.Subject = setting.Subject;
                    props.ExtendedProperties.GetUnderlyingProperties().Company = setting.Company;

                    return workbook;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        var workbook = new HSSFWorkbook(file);

                        _formulaEvaluator = new HSSFFormulaEvaluator(workbook);

                        return workbook;
                    }
                }
                else
                {
                    var workbook = new HSSFWorkbook();

                    _formulaEvaluator = new HSSFFormulaEvaluator(workbook);

                    var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                    dsi.Company = setting.Company;
                    workbook.DocumentSummaryInformation = dsi;

                    var si = PropertySetFactory.CreateSummaryInformation();
                    si.Author = setting.Author;
                    si.Subject = setting.Subject;
                    workbook.SummaryInformation = si;

                    return workbook;
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
