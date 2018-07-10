// Copyright (c) rigofunc (xuyingting). All rights reserved

namespace FluentExcel
{
    using System;

    /// <summary>
    /// Represents the configuration for the specfidied property.
    /// </summary>
    public class PropertyConfiguration
    {
        /// <summary>
        /// Gets the title of the excel column.
        /// </summary>
        /// <remarks>
        /// If the <see cref="Title"/> is null or empty, will use property name as the excel column title.
        /// </remarks>
        public string Title { get; internal set; }

        /// <summary>
        /// If <see cref="Index"/> was not set and AutoIndex is true FluentExcel will try to autodiscover the excel column index by its <see cref="Title"/> property.
        /// </summary>
        public bool AutoIndex { get; internal set; }

        /// <summary>
        /// Gets the exel column index.
        /// </summary>
        /// <value>The index.</value>
        public int Index { get; internal set; } = -1;

        /// <summary>
        /// Gets a value indicating whether allow merge the same value exel cells.
        /// </summary>
        public bool AllowMerge { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether this value of the property is ignored when exporting.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsExportIgnored { get; internal set; }

        /// <summary>
        /// Gets a value indicating whether this value of the property is ignored when importing.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsImportIgnored { get; internal set; }

        /// <summary>
        /// Gets the formatter for formatting the value.
        /// </summary>
        /// <value>The formatter.</value>
        public string Formatter { get; internal set; }

        /// <summary>
        /// Gets the cell value validator to validate the cell value.
        /// </summary>
        public CellValueValidator CellValueValidator { get; internal set; }

        /// <summary>
        /// Gets the value converter to convert the value.
        /// </summary>
        public CellValueConverter CellValueConverter { get; internal set; }

        /// <summary>
        /// Configures the excel cell index for the property.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="index">The excel cell index.</param>
        /// <remarks>
        /// If index was not set and AutoIndex is true FluentExcel will try to autodiscover the column index by its title setting.
        /// </remarks>
        public PropertyConfiguration HasExcelIndex(int index)
        {
            if (index < 0)
            {
                throw new IndexOutOfRangeException("The index cannot be less then 0");
            }

            Index = index;
            AutoIndex = false;

            return this;
        }

        /// <summary>
        /// Configures the excel title (first row) for the property.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="title">The excel cell title (fist row).</param>
        /// <remarks>
        /// If the title is string.Empty, will not set the excel cell, and if the title is NULL, the property's name will be used.
        /// </remarks>
        public PropertyConfiguration HasExcelTitle(string title)
        {
            Title = title;

            return this;
        }

        /// <summary>
        /// Configures the formatter will be used for formatting the value for the property.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="formatter">The formatter will be used for formatting the value.</param>
        /// <remarks>
        /// If the title is string.Empty, will not set the excel cell, and if the title is NULL, the property's name will be used.
        /// </remarks>
        public PropertyConfiguration HasDataFormatter(string formatter)
        {
            Formatter = formatter;

            return this;
        }

        /// <summary>
        /// Configures whether to autodiscover the column index by its title setting for the specified property.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <remarks>
        /// If index was not set and AutoIndex is true FluentExcel will try to autodiscover the column index by its title setting.
        /// </remarks>
        public PropertyConfiguration HasAutoIndex()
        {
            AutoIndex = true;
            Index = -1;

            return this;
        }

        /// <summary>
        /// Configures the value converter for the specified property.
        /// </summary>
        /// <param name="cellValueConverter">The value converter.</param>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        public PropertyConfiguration HasValueConverter(CellValueConverter cellValueConverter)
        {
            CellValueConverter = cellValueConverter;

            return this;
        }

        /// <summary>
        /// Configures the cell value validator for the specified property.
        /// </summary>
        /// <param name="cellValueValidator">The value validator.</param>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        public PropertyConfiguration HasValueValidator(CellValueValidator cellValueValidator)
        {
            CellValueValidator = cellValueValidator;

            return this;
        }

        /// <summary>
        /// Configures whether to allow merge the same value cells for the specified property.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        public PropertyConfiguration IsMergeEnabled()
        {
            AllowMerge = true;

            return this;
        }

        /// <summary>
        /// Configures whether to ignore the specified property when exporting or importing.
        /// </summary>
        /// <param name="exportingIsIgnored">If set to <c>true</c> exporting is ignored.</param>
        /// <param name="importingIsIgnored">If set to <c>true</c> importing is ignored.</param>
        public PropertyConfiguration IsIgnored(bool exportingIsIgnored, bool importingIsIgnored)
        {
            IsExportIgnored = exportingIsIgnored;
            IsImportIgnored = importingIsIgnored;

            return this;
        }

        /// <summary>
        /// Configures whether to ignore the specified property when exporting or importing.
        /// </summary>
        /// <param name="index">The excel cell index.</param>
        /// <param name="title">The excel cell title (fist row).</param>
        /// <param name="formatter">The formatter will be used for formatting the value.</param>
        /// <param name="exportingIsIgnored">If set to <c>true</c> exporting is ignored.</param>
        /// <param name="importingIsIgnored">If set to <c>true</c> importing is ignored.</param>
        public void IsIgnored(int index, string title, string formatter = null, bool exportingIsIgnored = true, bool importingIsIgnored = true)
        {
            if (index < 0)
            {
                throw new IndexOutOfRangeException("The index cannot be less then 0");
            }

            Index = index;
            Title = title;
            Formatter = formatter;
            IsExportIgnored = exportingIsIgnored;
            IsImportIgnored = importingIsIgnored;
        }

        /// <summary>
        /// Configures the excel cell for the property.
        /// </summary>
        /// <param name="index">The excel cell index.</param>
        /// <param name="title">The excel cell title (fist row).</param>
        /// <param name="formatter">The formatter will be used for formatting the value.</param>
        /// <param name="allowMerge">If set to <c>true</c> allow merge the same value cells.</param>
        /// <param name="cellValueConverter">The value converter.</param>
        public void HasExcelCell(int index, string title, string formatter = null, bool allowMerge = false, CellValueConverter cellValueConverter = null)
        {
            if (index < 0)
            {
                throw new IndexOutOfRangeException("The index cannot be less then 0");
            }

            Index = index;
            Title = title;
            Formatter = formatter;
            AutoIndex = false;
            AllowMerge = allowMerge;
            CellValueConverter = cellValueConverter;
        }

        /// <summary>
        /// Configures the excel cell for the property with index autodiscover. This method will try to autodiscover the column index by its <paramref name="title"/>
        /// </summary>
        /// <param name="title">The excel cell title (fist row).</param>
        /// <param name="formatter">The formatter will be used for formatting the value.</param>
        /// <param name="allowMerge">If set to <c>true</c> allow merge the same value cells.</param>
        /// <remarks>
        /// This method will try to autodiscover the column index by its <paramref name="title"/>
        /// </remarks>
        /// <param name="cellValueConverter">The value converter.</param>
        public void HasAutoIndexExcelCell(string title, string formatter = null, bool allowMerge = false, CellValueConverter cellValueConverter = null)
        {
            Index = -1;
            Title = title;
            Formatter = formatter;
            AutoIndex = true;
            AllowMerge = allowMerge;
            CellValueConverter = cellValueConverter;
        }
    }
}
