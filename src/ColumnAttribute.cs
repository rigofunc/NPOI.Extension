// Copyright (c) RigoFunc (xuyingting). All rights reserved

namespace NPOI.Extension {
    using System;

    /// <summary>
    /// Represents a custom attribute to control object's properties to excel columns behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColumnAttribute : Attribute {
        /// <summary>
        /// Gets or sets the title of the column.
        /// </summary>
        /// <remarks>
        /// If the <see cref="Title"/> is null or empty, will use property name as the excel column title.
        /// </remarks>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the index of the column.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether allow merge the same value cells.
        /// </summary>
        public bool AllowMerge { get; set; }
    }
}
