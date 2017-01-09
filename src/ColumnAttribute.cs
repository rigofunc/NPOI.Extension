// Copyright (c) RigoFunc (xuyingting). All rights reserved

namespace NPOI.Extension
{
    using System;

    /// <summary>
    /// Represents a custom attribute to control object's properties to excel columns behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColumnAttribute : Attribute
    {
        private int index = -1;

        /// <summary>
        /// Gets or sets the title of the column.
        /// </summary>
        /// <remarks>
        /// If the <see cref="Title"/> is null or empty, will use property name as the excel column title.
        /// </remarks>
        public string Title { get; set; }

        /// <summary>
        /// If <see cref="Index"/> was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its <see cref="Title"/> property.
        /// </summary>
        public bool AutoIndex { get; set; }

        public int Index
        {
            get { return index; }
            set { index = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow merge the same value cells.
        /// </summary>
        public bool AllowMerge { get; set; }
    }
}
