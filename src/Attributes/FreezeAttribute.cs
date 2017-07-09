// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System;

    /// <summary>
    /// Represents a custom attribute to control excel freeze behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class FreezeAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FreezeAttribute"/> class.
        /// </summary>
        public FreezeAttribute()
        {
            FreezeConfig = new FreezeConfig();
        }

        /// <summary>
        /// Gets or sets the column number to split.
        /// </summary>
        public int ColSplit
        {
            get
            {
                return FreezeConfig.ColSplit;
            }
            set
            {
                FreezeConfig.ColSplit = value;
            }
        }

        /// <summary>
        /// Gets or sets the row number to split.
        /// </summary>
        public int RowSplit
        {
            get
            {
                return FreezeConfig.RowSplit;
            }
            set
            {
                FreezeConfig.RowSplit = value;
            }
        }

        /// <summary>
        /// Gets or sets the left most culomn index.
        /// </summary>
        public int LeftMostColumn
        {
            get
            {
                return FreezeConfig.LeftMostColumn;
            }
            set
            {
                FreezeConfig.LeftMostColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets the top most row index.
        /// </summary>
        public int TopRow
        {
            get
            {
                return FreezeConfig.TopRow;
            }
            set
            {
                FreezeConfig.TopRow = value;
            }
        }

        /// <summary>
        /// Gets the freeze config.
        /// </summary>
        /// <value>The freeze config.</value>
        internal FreezeConfig FreezeConfig { get; }
    }
}
