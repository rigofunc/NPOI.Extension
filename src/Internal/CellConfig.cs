// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace NPOI.Extension
{
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
		/// If <see cref="Index"/> was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its <see cref="Title"/> property.
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
        public bool AllowMerge { get; set; } = true;
    }
}
