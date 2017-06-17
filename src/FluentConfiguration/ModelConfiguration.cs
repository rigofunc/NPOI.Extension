// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace NPOI.Extension
{
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

	/// <summary>
	/// Represents the configuration for the specfidied <typeparamref name="TModel"/>.
	/// </summary>
	/// <typeparam name="TModel">The type of model.</typeparam>
	public class ModelConfiguration<TModel> where TModel : class
    {
		/// <summary>
		/// Gets the property configuration by the specified property expression.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration{TModel, TProperty}"/>.</returns>
		/// <param name="propertyExpression">The property expression.</param>
		/// <typeparam name="TProperty">The type of parameter.</typeparam>
		public PropertyConfiguration<TModel, TProperty> Property<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            return new PropertyConfiguration<TModel, TProperty>(propertyExpression);
        }

		/// <summary>
		/// Configures the statistics for the specified <typeparamref name="TModel"/>. Only for vertical, not for horizontal statistics.
		/// </summary>
		/// <returns>The <see cref="ModelConfiguration{TModel}"/>.</returns>
		/// <param name="name">The statistics name. (e.g. Total). In current version, the default name location is (last row, first cell)</param>
		/// <param name="formula">The cell formula, such as SUM, AVERAGE and so on, which applyable for vertical statistics..</param>
		/// <param name="columnIndexes">The column indexes for statistics. if <paramref name="formula"/>is SUM, and <paramref name="columnIndexes"/> is [1,3], 
		/// for example, the column No. 1 and 3 will be SUM for first row to last row.</param>
		public ModelConfiguration<TModel> HasStatistics(string name, string formula, params int[] columnIndexes)
        {
            return this;
        }

		/// <summary>
		/// Configures the excel filter behaviors for the specified <typeparamref name="TModel"/>.
		/// </summary>
		/// <returns>The <see cref="ModelConfiguration{TModel}"/>.</returns>
		/// <param name="firstColumn">The first column index.</param>
		/// <param name="lastColumn">The last column index.</param>
		/// <param name="firstRow">The first row index.</param>
		/// <param name="lastRow">The last row index. If is null, the value is dynamic calculate by code.</param>
		public ModelConfiguration<TModel> HasFilter(int firstColumn, int lastColumn, int firstRow, int? lastRow = null)
        {
            return this;
        }

		/// <summary>
		/// Configures the excel freeze behaviors for the specified <typeparamref name="TModel"/>.
		/// </summary>
		/// <returns>The <see cref="ModelConfiguration{TModel}"/>.</returns>
		/// <param name="columnSplit">The column number to split.</param>
		/// <param name="rowSplit">The row number to split.param>
		/// <param name="leftMostColumn">The left most culomn index.</param>
		/// <param name="topMostRow">The top most row index.</param>
		public ModelConfiguration<TModel> HasFreeze(int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
		{
			return this;
		}
    }
}