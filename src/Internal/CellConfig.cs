// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    /// <summary>
    /// Represents the excel cell configuration for the specified model's property.
    /// </summary>
    internal class CellConfig
    {
        /// <summary>
        /// Gets or sets the title of the column.
        /// </summary>
        /// <remarks>
        /// If the <see cref="Title"/> is null or empty, will use property name as the excel column title.
        /// </remarks>
        public string Title { get; set; }

        /// <summary>
        /// If <see cref="Index"/> was not set and AutoIndex is true FluentExcel will try to autodiscover the column index by its <see cref="Title"/> property.
        /// </summary>
        public bool AutoIndex { get; set; }

        /// <summary>
        /// Gets or sets the column index.
        /// </summary>
        /// <value>The index.</value>
        public int Index { get; set; } = -1;

        /// <summary>
        /// Gets or sets a value indicating whether allow merge the same value cells.
        /// </summary>
        public bool AllowMerge { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this value of the property is ignored when exporting.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsExportIgnored { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this value of the property is ignored when importing.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsImportIgnored { get; set; }

        /// <summary>
        /// Gets or sets the formatter for formatting the value.
        /// </summary>
        /// <value>The formatter.</value>
        public string Formatter { get; set; }
    }
}
