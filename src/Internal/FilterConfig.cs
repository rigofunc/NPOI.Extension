// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    /// <summary>
    /// Represents the excel fileter configration for the specified model.
    /// </summary>
    internal class FilterConfig
    {
        /// <summary>
        /// Gets or sets the first row index.
        /// </summary>
        public int FirstRow { get; set; }

        /// <summary>
        /// Gets or sets  the last row index.
        /// </summary>
        /// <remarks>
        /// If the <see cref="LastRow"/> is null, the value is dynamic calculate by code.
        /// </remarks>
        public int? LastRow { get; set; } = null;

        /// <summary>
        /// Gets or sets the first column index.
        /// </summary>
        public int FirstCol { get; set; }

        /// <summary>
        /// Gets or sets the last column index.
        /// </summary>
        public int LastCol { get; set; }
    }
}
