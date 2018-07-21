// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using NPOI.SS.UserModel;

    /// <summary>
    /// Provides some methods for loading <see cref="IEnumerable{T}"/> from excel.
    /// </summary>
    public static class Excel
    {
        private static IFormulaEvaluator _formulaEvaluator;

        /// <summary>
        /// Gets or sets the setting.
        /// </summary>
        /// <value>The setting.</value>
        public static ExcelSetting Setting { get; set; } = new ExcelSetting();

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, int startRow = 1, int sheetIndex = 0) where T : class, new()
            => Load<T>(excelFile, Setting, startRow, sheetIndex);

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="excelSetting">The excel setting to use to load data.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, ExcelSetting excelSetting, int startRow = 1, int sheetIndex = 0) where T : class, new()
        {
            if (!File.Exists(excelFile)) throw new FileNotFoundException();

            return Load<T>(File.OpenRead(excelFile), excelSetting, startRow, sheetIndex);
        }

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="sheetName">Which sheet to read.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, string sheetName, int startRow = 1) where T : class, new()
            => Load<T>(excelFile, Setting, sheetName, startRow);

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="excelSetting">The excel setting to use to load data.</param>
        /// <param name="sheetName">Which sheet to read.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, ExcelSetting excelSetting, string sheetName, int startRow = 1) where T : class, new()
        {
            if (!File.Exists(excelFile)) throw new FileNotFoundException();

            return Load<T>(File.OpenRead(excelFile), excelSetting, sheetName, startRow);
        }

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel stream.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelStream">The excel stream.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(Stream excelStream, int startRow = 1, int sheetIndex = 0) where T : class, new()
            => Load<T>(excelStream, Setting, startRow, sheetIndex);

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel stream.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelStream">The excel stream.</param>
        /// <param name="excelSetting">The excel setting to use to load data.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(Stream excelStream, ExcelSetting excelSetting, int startRow = 1, int sheetIndex = 0) where T : class, new()
        {
            var workbook = InitializeWorkbook(excelStream);

            // currently, only handle one sheet (or call side using foreach to support multiple sheet)
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (null == sheet) throw new ArgumentException($"Excel sheet with specified index [{sheetIndex}] not found", nameof(sheetIndex));

            return Load<T>(sheet, _formulaEvaluator, excelSetting, startRow);
        }

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel stream.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelStream">The excel stream.</param>
        /// <param name="sheetName">Which sheet to read.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(Stream excelStream, string sheetName, int startRow = 1) where T : class, new()
            => Load<T>(excelStream, Setting, sheetName, startRow);

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel stream.
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelStream">The excel stream.</param>
        /// <param name="excelSetting">The excel setting to use to load data.</param>
        /// <param name="sheetName">Which sheet to read.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(Stream excelStream, ExcelSetting excelSetting, string sheetName, int startRow = 1) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException($"sheet name cannot be null or whitespace", nameof(sheetName));

            var workbook = InitializeWorkbook(excelStream);

            // currently, only handle one sheet (or call side using foreach to support multiple sheet)
            var sheet = workbook.GetSheet(sheetName);
            if (null == sheet) throw new ArgumentException($"Excel sheet with specified name [{sheetName}] not found", nameof(sheetName));

            return Load<T>(sheet, _formulaEvaluator, excelSetting, startRow);
        }

        public static IEnumerable<T> Load<T>(ISheet sheet, IFormulaEvaluator formulaEvaluator, int startRow = 1) where T : class, new()
            => Load<T>(sheet, formulaEvaluator, Setting, startRow);

        public static IEnumerable<T> Load<T>(ISheet sheet, IFormulaEvaluator formulaEvaluator, ExcelSetting excelSetting, int startRow = 1) where T : class, new()
        {
            if (null == sheet) throw new ArgumentNullException(nameof(sheet));

            // get the writable properties
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);

            bool fluentConfigEnabled = false;
            // get the fluent config
            if (excelSetting.FluentConfigs.TryGetValue(typeof(T), out var fluentConfig))
            {
                fluentConfigEnabled = true;
            }

            var propertyConfigurations = new PropertyConfiguration[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];
                if (fluentConfigEnabled && fluentConfig.PropertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    // fluent configure first(Hight Priority)
                    propertyConfigurations[j] = pc;
                }
                else
                {
                    propertyConfigurations[j] = null;
                }
            }

            var statistics = new List<StatisticsConfiguration>();
            if (fluentConfigEnabled)
            {
                statistics.AddRange(fluentConfig.StatisticsConfigurations);
            }

            var list = new List<T>();
            int idx = 0;

            IRow headerRow = null;

            // get the physical rows
            var rows = sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                var row = rows.Current as IRow;

                if (idx == 0)
                    headerRow = row;
                idx++;

                if (row.RowNum < startRow)
                {
                    continue;
                }

                // ignore whitespace rows if requested
                if (true == fluentConfig?.IgnoreWhitespaceRows)
                {
                    if (row.Cells.All(x =>
                        CellType.Blank == x.CellType
                        || (CellType.String == x.CellType && string.IsNullOrWhiteSpace(x.StringCellValue))
                    )) continue;
                }

                var item = new T();
                var itemIsValid = true;
                for (int i = 0; i < properties.Length; i++)
                {
                    var prop = properties[i];

                    int index = i;
                    var config = propertyConfigurations[i];
                    if (config != null)
                    {
                        if (config.IsImportIgnored)
                            continue;

                        index = config.Index;

                        // Try to autodiscover index from title and cache
                        if (index < 0 && config.AutoIndex && !string.IsNullOrEmpty(config.Title))
                        {
                            foreach (var cell in headerRow.Cells)
                            {
                                if (!string.IsNullOrEmpty(cell.StringCellValue))
                                {
                                    if (cell.StringCellValue.Equals(config.Title, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        index = cell.ColumnIndex;

                                        // cache
                                        config.Index = index;

                                        break;
                                    }
                                }
                            }
                        }

                        // check again
                        if (index < 0)
                        {
                            throw new ApplicationException("Please set the 'index' or 'autoIndex' by fluent api or attributes");
                        }
                    }

                    var value = row.GetCellValue(index, formulaEvaluator);

                    // give a chance to the cell value validator
                    if (null != config?.CellValueValidator)
                    {
                        var validationResult = config.CellValueValidator(row.RowNum - 1, config.Index, value);
                        if (false == validationResult)
                        {
                            if (fluentConfig.SkipInvalidRows)
                            {
                                itemIsValid = false;
                                break;
                            }

                            throw new ArgumentException($"Validation of cell value at row {row.RowNum}, column {config.Title}({config.Index + 1}) failed! Value: [{value}]");
                        }
                    }

                    // give a chance to the value converter.
                    if (config?.CellValueConverter != null)
                    {
                        value = config.CellValueConverter(row.RowNum - 1, config.Index, value);
                    }

                    if (value == null)
                    {
                        continue;
                    }

                    // check whether is statics row
                    if (idx > startRow + 1 && index == 0
                        &&
                        statistics.Any(s => s.Name.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase)))
                    {
                        var st = statistics.FirstOrDefault(s => s.Name.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
                        var formula = row.GetCellValue(st.Columns.First()).ToString();
                        if (formula.StartsWith(st.Formula, StringComparison.InvariantCultureIgnoreCase))
                        {
                            itemIsValid = false;
                            break;
                        }
                    }

                    // property type
                    var propType = prop.PropertyType.UnwrapNullableType();

                    var safeValue = Convert.ChangeType(value, propType, CultureInfo.CurrentCulture);

                    prop.SetValue(item, safeValue, null);
                }

                if (itemIsValid)
                {
                    // give a chance to the row data validator
                    if (null != fluentConfig?.RowDataValidator)
                    {
                        var validationResult = fluentConfig.RowDataValidator(row.RowNum - 1, item);
                        if (false == validationResult)
                        {
                            if (fluentConfig.SkipInvalidRows)
                            {
                                itemIsValid = false;
                                continue;
                            }

                            throw new ArgumentException($"Validation of row data at row {row.RowNum} failed!");
                        }
                    }

                    list.Add(item);
                }
            }

            return list;
        }

        internal static object GetCellValue(this IRow row, int index, IFormulaEvaluator eval = null)
        {
            var cell = row.GetCell(index);
            if (cell == null)
            {
                return null;
            }

            return cell.GetCellValue(eval);
        }

        internal static object GetCellValue(this ICell cell, IFormulaEvaluator eval = null)
        {
            if (cell.IsMergedCell)
            {
                // what can I do here?
            }

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Error:
                    return FormulaError.ForInt(cell.ErrorCellValue).String;

                case CellType.Formula:
                    if (eval != null)
                        return GetCellValue(eval.EvaluateInCell(cell));
                    else
                        return cell.CellFormula;

                case CellType.Blank:
                case CellType.Unknown:
                default:
                    return null;
            }
        }

        private static IWorkbook InitializeWorkbook(string excelFile)
            => InitializeWorkbook(File.OpenRead(excelFile));

        private static IWorkbook InitializeWorkbook(Stream excelStream)
        {
            var workbook = WorkbookFactory.Create(excelStream);
            _formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            return workbook;
        }
    }
}
