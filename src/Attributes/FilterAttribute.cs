// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System;

    /// <summary>
    /// Represents a custom attribute to control excel filter behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class FilterAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FilterAttribute"/> class.
        /// </summary>
        public FilterAttribute()
        {
            FilterConfig = new FilterConfig();
        }

        /// <summary>
        /// Gets or sets the first row index.
        /// </summary>
        public int FirstRow
        {
            get
            {
                return FilterConfig.FirstRow;
            }
            set
            {
                FilterConfig.FirstRow = value;
            }
        }

        /// <summary>
        /// Gets or sets  the last row index.
        /// </summary>
        /// <remarks>
        /// If the <see cref="LastRow"/> is null, the value is dynamic calculate by code.
        /// </remarks>
        public int? LastRow
        {
            get
            {
                return FilterConfig.LastRow;
            }
            set
            {
                FilterConfig.LastRow = value;
            }
        }

        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstCol
        {
            get
            {
                return FilterConfig.FirstCol;
            }
            set
            {
                FilterConfig.FirstCol = value;
            }
        }

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int LastCol
        {
            get
            {
                return FilterConfig.LastCol;
            }
            set
            {
                FilterConfig.LastCol = value;
            }
        }

        /// <summary>
        /// Gets or the filter config.
        /// </summary>
        /// <value>The filter config.</value>
        internal FilterConfig FilterConfig { get; }
    }
}
