// Copyright (c) rigofunc (xuyingting). All rights reserved

namespace NPOI.Extension
{
	/// <summary>
	/// Represents the configuration for the specfidied property.
	/// </summary>
	public class PropertyConfiguration
    {
		/// <summary>
        /// Initializes a new instance of the <see cref="PropertyConfiguration"/> class.
		/// </summary>
		public PropertyConfiguration()
        {
            CellConfig = new CellConfig();
        }

        /// <summary>
        /// Gets the cell config.
        /// </summary>
        /// <value>The cell config.</value>
        internal CellConfig CellConfig { get; }

		/// <summary>
		/// Configures the excel cell index for the property.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration"/>.</returns>
		/// <param name="index">The excel cell index.</param>
		/// <remarks>
		/// If index was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its title setting.
		/// </remarks>
		public PropertyConfiguration HasExcelIndex(int index)
        {
            CellConfig.Index = index;

            return this;
        }

		/// <summary>
		/// Configures the excel title (first row) for the property.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration"/>.</returns>
		/// <param name="title">The excel cell title (fist row).</param>
		/// <remarks>
		/// If the title is string.Empty, will not set the excel cell, and if the title is NULL, the property's name will be used.
		/// </remarks>
		public PropertyConfiguration HasExcelTitle(string title) 
        {
            CellConfig.Title = title;

            return this;
        }

		/// <summary>
		/// Configures whether to autodiscover the column index by its title setting for the specified property.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration"/>.</returns>
		/// <remarks>
		/// If index was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its title setting.
		/// </remarks>
		public PropertyConfiguration HasAutoIndex()
        {
            CellConfig.AutoIndex = true;

            return this;
        }

		/// <summary>
		/// Configures whether to allow merge the same value cells for the specified property.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration"/>.</returns>
		public PropertyConfiguration IsMergeEnabled()
        {
            CellConfig.AllowMerge = true;

            return this;
        }

		/// <summary>
		/// Configures the excel cell for the property.
		/// </summary>
		/// <param name="index">The excel cell index.</param>
		/// <param name="title">The excel cell title (fist row).</param>
		/// <param name="allowMerge">If set to <c>true</c> allow merge the same value cells.</param>
		public void HasExcelCell(int index, string title, bool allowMerge)
        {
            CellConfig.Index = index;
            CellConfig.Title = title;
            CellConfig.AutoIndex = false;
            CellConfig.AllowMerge = allowMerge;
        }

		/// <summary>
		/// Configures the excel cell for the property.
		/// </summary>
		/// <param name="title">The excel cell title (fist row).</param>
		/// <param name="allowMerge">If set to <c>true</c> allow merge the same value cells.</param>
		/// <remarks>
		/// This method will try to autodiscover the column index by its <paramref name="title"/>
		/// </remarks>
		public void HasExcelCell(string title, bool allowMerge)
		{
			CellConfig.Index = -1;
			CellConfig.Title = title;
            CellConfig.AutoIndex = true;
			CellConfig.AllowMerge = allowMerge;
		}
    }
}
