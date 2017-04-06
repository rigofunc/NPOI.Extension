// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace NPOI.Extension
{
    using System;

    /// <summary>
    /// Represents a custom attribute to control excel freeze behaviors.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class FreezeAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the column number to split.
        /// </summary>
        public int ColSplit { get; set; } = 0;

        /// <summary>
        /// Gets or sets the row number to split.
        /// </summary>
        public int RowSplit { get; set; } = 1;

        /// <summary>
        /// Gets or sets the left most culomn index.
        /// </summary>
        public int LeftMostColumn { get; set; } = 0;

        /// <summary>
        /// Gets or sets the top most row index.
        /// </summary>
        public int TopRow { get; set; } = 1;
    }
}
