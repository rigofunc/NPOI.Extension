// Copyright (c) RigoFunc (xuyingting). All rights reserved.

namespace NPOI.Extension {
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Reflection;
    using HSSF.UserModel;
    using SS.UserModel;

    public delegate object ValueConverter(int row, int cell, object value);

    public static class Excel {
        public static IEnumerable<T> Load<T>(string excelFile, int startRow = 1, ValueConverter valueConverter = null) where T : new() {
            if (!File.Exists(excelFile)) {
                throw new FileNotFoundException();
            }

            var workbook = InitializeWorkbook(excelFile);

            // currently, only handle sheet one
            var sheet = workbook.GetSheetAt(0);

            // get the physical rows
            var rows = sheet.GetRowEnumerator();

            // get the writable properties
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);

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

            var list = new List<T>();
            while (rows.MoveNext()) {
                var row = rows.Current as HSSFRow;

                if (row.RowNum < startRow) {
                    continue;
                }

                var item = new T();
                for (int i = 0; i < properties.Length; i++) {
                    var prop = properties[i];
                    var attr = attributes[i];
                    if (attr == null) {
                        continue;
                    }

                    var value = row.GetCellValue(attr.Index);
                    if (valueConverter != null) {
                        value = valueConverter(row.RowNum, attr.Index, value);
                    }
                    if (value != null) {
                        // property type
                        var propType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                        var safeValue = Convert.ChangeType(value, propType, CultureInfo.CurrentCulture);

                        prop.SetValue(item, safeValue, null);
                    }
                }

                list.Add(item);
            }

            return list;
        }

        internal static object GetCellValue(this IRow row, int index) {
            var cell = row.GetCell(index);
            if (cell == null) {
                return null;
            }

            switch (cell.CellType) {
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

        internal static object GetDefault(this Type type) {
            if (type.IsValueType) {
                return Activator.CreateInstance(type);
            }

            return null;
        }

        private static HSSFWorkbook InitializeWorkbook(string excelFile) {
            using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read)) {
                return new HSSFWorkbook(file);
            }
        }
    }
}
