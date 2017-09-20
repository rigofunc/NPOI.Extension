// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    /// <summary>
    /// Represents the excel freeze configuration for the specified model.
    /// </summary>
    public class FreezeConfiguration
    {
        /// <summary>
        /// Gets the column number to split.
        /// </summary>
        public int ColSplit { get; internal set; } = 0;

        /// <summary>
        /// Gets the row number to split.
        /// </summary>
        public int RowSplit { get; internal set; } = 1;

        /// <summary>
        /// Gets the left most culomn index.
        /// </summary>
        public int LeftMostColumn { get; internal set; } = 0;

        /// <summary>
        /// Gets the top most row index.
        /// </summary>
        public int TopRow { get; internal set; } = 1;
    }
}
