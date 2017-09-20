// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    /// <summary>
    /// Represents the excel fileter configration for the specified model.
    /// </summary>
    public class FilterConfiguration
    {
        /// <summary>
        /// Gets the first row index.
        /// </summary>
        public int FirstRow { get; internal set; }

        /// <summary>
        /// Gets the last row index.
        /// </summary>
        /// <remarks>
        /// If the <see cref="LastRow"/> is null, the value is dynamic calculate by code.
        /// </remarks>
        public int? LastRow { get; internal set; } = null;

        /// <summary>
        /// Gets the first column index.
        /// </summary>
        public int FirstCol { get; internal set; }

        /// <summary>
        /// Gets the last column index.
        /// </summary>
        public int LastCol { get; internal set; }
    }
}
