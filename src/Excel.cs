// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    /// <summary>
    /// Represents the cell value converter, which convert the value to another value.
    /// </summary>
    /// <param name="row">The row of the excel sheet.</param>
    /// <param name="cell">The cell of the excel sheet.</param>
    /// <param name="value">The value to convert.</param>
    /// <returns>The converted value.</returns>
    public delegate object ValueConverter(int row, int cell, object value);

    /// <summary>
    /// Provides some methods for loading <see cref="IEnumerable{T}"/> from excel.
    /// </summary>
    public static class Excel
    {
        /// <summary>
        /// Gets or sets the setting.
        /// </summary>
        /// <value>The setting.</value>
        public static ExcelSetting Setting { get; set; } = new ExcelSetting();

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file.
        /// /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <param name="valueConverter">The cell value convert.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, int startRow = 1, int sheetIndex = 0, ValueConverter valueConverter = null) where T : class, new()
        {
            if (!File.Exists(excelFile))
            {
                throw new FileNotFoundException();
            }

            var workbook = InitializeWorkbook(excelFile);

            // currently, only handle sheet one (or call side using foreach to support multiple sheet)
            var sheet = workbook.GetSheetAt(sheetIndex);

            // get the writable properties
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);

            bool fluentConfigEnabled = false;
            // get the fluent config
            if (Excel.Setting.FluentConfigs.TryGetValue(typeof(T), out var fluentConfig))
            {
                fluentConfigEnabled = true;
            }

            var cellConfigs = new CellConfig[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];
                if (fluentConfigEnabled && fluentConfig.PropertyConfigs.TryGetValue(property, out var pc))
                {
                    // fluent configure first(Hight Priority)
                    cellConfigs[j] = pc.CellConfig;
                }
                else
                {
                    var attrs = property.GetCustomAttributes(typeof(ColumnAttribute), true) as ColumnAttribute[];
                    if (attrs != null && attrs.Length > 0)
                    {
                        cellConfigs[j] = attrs[0].CellConfig;
                    }
                    else
                    {
                        cellConfigs[j] = null;
                    }
                }
            }

            var statistics = new List<StatisticsConfig>();
            if (fluentConfigEnabled)
            {
                statistics.AddRange(fluentConfig.StatisticsConfigs);
            }
            else
            {
                var attributes = typeof(T).GetCustomAttributes(typeof(StatisticsAttribute), true) as StatisticsAttribute[];
                if (attributes != null && attributes.Length > 0)
                {
                    foreach (var item in attributes)
                    {
                        statistics.Add(item.StatisticsConfig);
                    }
                }
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

                var item = new T();
                var itemIsValid = true;
                for (int i = 0; i < properties.Length; i++)
                {
                    var prop = properties[i];

                    int index = i;
                    var config = cellConfigs[i];
                    if (config != null)
                    {
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

                    var value = row.GetCellValue(index);
                    if (valueConverter != null)
                    {
                        value = valueConverter(row.RowNum, index, value);
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
                    var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                    var safeValue = Convert.ChangeType(value, propType, CultureInfo.CurrentCulture);

                    prop.SetValue(item, safeValue, null);
                }

                if (itemIsValid)
                {
                    list.Add(item);
                }
            }

            return list;
        }

        internal static object GetCellValue(this IRow row, int index)
        {
            var cell = row.GetCell(index);
            if (cell == null)
            {
                return null;
            }

            if (cell.IsMergedCell)
            {
                // what can I do here?

            }

            switch (cell.CellType)
            {
                // This is a trick to get the correct value of the cell.
                // NumericCellValue will return a numeric value no matter the cell value is a date or a number.
                case CellType.Numeric:
                    return cell.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;

                // how?
                case CellType.Formula:
                    return cell.ToString();

                case CellType.Blank:
                case CellType.Unknown:
                default:
                    return null;
            }
        }

        internal static object GetDefault(this Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }

        private static IWorkbook InitializeWorkbook(string excelFile)
        {
            if (Path.GetExtension(excelFile).Equals(".xls"))
            {
                using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                {
                    return new HSSFWorkbook(file);
                }
            }
            else if (Path.GetExtension(excelFile).Equals(".xlsx"))
            {
                using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                {
                    return new XSSFWorkbook(file);
                }
            }
            else
            {
                throw new NotSupportedException($"not an excel file {excelFile}");
            }
        }
    }
}
