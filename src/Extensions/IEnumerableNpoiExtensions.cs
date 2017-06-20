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

            bool fluentConfigEnabled = false;
            // get the fluent config
            if (Excel.Setting.FluentConfigs.TryGetValue(typeof(T), out var fluentConfig))
            {
                fluentConfigEnabled = true;
            }

            // find out the configs
            var haventCols = true;
            var cellConfigs = new CellConfig[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];

                // get the property config
                if (fluentConfigEnabled && fluentConfig.PropertyConfigs.TryGetValue(property, out var pc))
                {
					// fluent configure first(Hight Priority)
                    cellConfigs[j] = pc.CellConfig;
					haventCols = false;
                }
                else
                {
					var attrs = property.GetCustomAttributes(typeof(ColumnAttribute), true) as ColumnAttribute[];
					if (attrs != null && attrs.Length > 0)
					{
                        cellConfigs[j] = attrs[0].CellConfig;
						haventCols = false;
					}
					else
					{
						cellConfigs[j] = null;
					}
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
                        var config = cellConfigs[i];
                        if (config == null)
                            continue;
                        else
                            index = config.Index;
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
                for (var j = 0; j < cellConfigs.Length; j++)
                {
                    var config = cellConfigs[j];
                    if (config == null)
                    {
                        continue;
                    }

					var previous = "";
                    //object previous = null;
                    int rowspan = 0, row = 1;
                    if (config.AllowMerge)
                    {
                        for (row = 1; row < rowIndex; row++)
                        {
							var value = sheet.GetRow(row).Cells[config.Index].StringCellValue;
							if (previous == value && !string.IsNullOrEmpty(value))
                            //var value = sheet.GetRow(row).GetCellValue(config.Index);
                            //if (previous == value && value != null)
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
                    var config = cellConfigs[i];
                    if (config == null)
                        continue;
                    else
                    {
                        index = config.Index;
                        // if not title, using property name as title.
                        if (!string.IsNullOrEmpty(config.Title))
                        {
                            title = config.Title;
                        }
                    }
                }

                var cell = row1.CreateCell(index);
                cell.CellStyle = style;
                cell.SetCellValue(title);
            }

            if (rowIndex > 0)
            {
                var statistics = new List<StatisticsConfig>();
		        var filterConfigs = new List<FilterConfig>();
		        var freezeConfigs = new List<FreezeConfig>();
                if (fluentConfigEnabled) 
                {
                    statistics.AddRange(fluentConfig.StatisticsConfigs);
                    freezeConfigs.AddRange(fluentConfig.FreezeConfigs);
                    filterConfigs.AddRange(fluentConfig.FilterConfigs);
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

					var freezes = typeof(T).GetCustomAttributes(typeof(FreezeAttribute), true) as FreezeAttribute[];
					if (freezes != null && freezes.Length > 0)
					{
                        foreach (var item in freezes)
                        {
                            freezeConfigs.Add(item.FreezeConfig);
                        }
                    }

					var filters = typeof(T).GetCustomAttributes(typeof(FilterAttribute), true) as FilterAttribute[];
					if (filters != null && filters.Length > 0)
					{
                        foreach (var item in filters)
                        {
                            filterConfigs.Add(item.FilterConfig);
                        }
                    }
                }

				// statistics row
				foreach (var item in statistics)
				{
					var lastRow = sheet.CreateRow(rowIndex);
					var cell = lastRow.CreateCell(0);
					cell.SetCellValue(item.Name);
					foreach (var column in item.Columns)
					{
						cell = lastRow.CreateCell(column);
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
