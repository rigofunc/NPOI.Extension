// Copyright (c) rigofunc (xuyingting). All rights reserved

namespace Arch.FluentExcel
{
    using System;

    /// <summary>
    /// Represents a custom attribute to control object's properties to excel columns behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColumnAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnAttribute"/> class.
        /// </summary>
        public ColumnAttribute()
        {
            CellConfig = new CellConfig();
        }

        /// <summary>
        /// Gets or sets the title of the column.
        /// </summary>
        /// <remarks>
        /// If the <see cref="Title"/> is null or empty, will use property name as the excel column title.
        /// </remarks>
        public string Title
        {
            get
            {
                return CellConfig.Title;
            }
            set
            {
                CellConfig.Title = value;
            }
        }

        /// <summary>
        /// If <see cref="Index"/> was not set and AutoIndex is true Arch.FluentExcel will try to autodiscover the column index by its <see cref="Title"/> property.
        /// </summary>
        public bool AutoIndex
        {
            get
            {
                return CellConfig.AutoIndex;
            }
            set
            {
                CellConfig.AutoIndex = value;
            }
        }

        /// <summary>
        /// Gets or sets the column index.
        /// </summary>
        /// <value>The index.</value>
        public int Index
        {
            get
            {
                return CellConfig.Index;
            }
            set
            {
                CellConfig.Index = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow merge the same value cells.
        /// </summary>
        public bool AllowMerge
        {
            get
            {
                return CellConfig.AllowMerge;
            }
            set
            {
                CellConfig.AllowMerge = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this value of the propery is ignored.
        /// </summary>
        /// <value><c>true</c> if is ignored; otherwise, <c>false</c>.</value>
        public bool IsIgnored
        {
            get
            {
                return CellConfig.IsIgnored;
            }
            set
            {
                CellConfig.IsIgnored = value;
            }
        }

        /// <summary>
        /// Gets or sets the formatter for formatting the value.
        /// </summary>
        /// <value>The format.</value>
		public string Formatter
        {
            get
            {
                return CellConfig.Formatter;
            }
            set
            {
                CellConfig.Formatter = value;
            }
        }

        /// <summary>
        /// Gets the cell config.
        /// </summary>
        /// <value>The cell config.</value>
        internal CellConfig CellConfig { get; }
    }
}
