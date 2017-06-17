// Copyright (c) rigofunc (xuyingting). All rights reserved

namespace NPOI.Extension
{
    using System;
    using System.Linq.Expressions;

	/// <summary>
	/// Represents the configuration for the specfidied <typeparam name="TProperty"> of the specified <typeparamref name="TModel"/>.
	/// </summary>
	/// <typeparam name="TModel">The type of model.</typeparam>
	/// <typeparam name="TProperty">The type of property.</typeparam>
	public class PropertyConfiguration<TModel, TProperty> where TModel : class
    {
		/// <summary>
        /// Initializes a new instance of the <see cref="PropertyConfiguration{TModel, TProperty}"/> class.
		/// </summary>
		/// <param name="propertyExpression">The property expression.</param>
		public PropertyConfiguration(Expression<Func<TModel, TProperty>> propertyExpression)
        {
        }

		/// <summary>
        /// Configures the excel cell index for the <typeparamref name="TProperty"/>.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration{TModel, TProperty}"/>.</returns>
		/// <param name="index">The excel cell index.</param>
		/// <remarks>
		/// If index was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its title setting.
		/// </remarks>
		public PropertyConfiguration<TModel, TProperty> HasExcelIndex(int index)
        {
            return this;
        }

		/// <summary>
        /// Configures the excel title (first row) for the <typeparamref name="TProperty"/>.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration{TModel, TProperty}"/>.</returns>
		/// <param name="title">Title.</param>
        /// <remarks>
        /// If the title is string.Empty, will not set the excel cell, and if the title is NULL, the property's name will be used.
        /// </remarks>
		public PropertyConfiguration<TModel, TProperty> HasExcelTitle(string title) 
        {
            return this;
        }

		/// <summary>
        /// Configures whether to autodiscover the column index by its title setting for the specified <typeparamref name="TProperty"/>.
		/// </summary>
		/// <returns>The <see cref="PropertyConfiguration{TModel, TProperty}"/>.</returns>
		/// <param name="autoIndex">If set to <c>true</c> auto index.</param>
		/// <remarks>
		/// If index was not set and AutoIndex is true NPOI.Extension will try to autodiscover the column index by its title setting.
		/// </remarks>
		public PropertyConfiguration<TModel, TProperty> HasAutoIndex(bool autoIndex)
        {
            return this;
        }

        /// <summary>
        /// Configures whether to allow merge the same value cells for the specified <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration{TModel, TProperty}"/>.</returns>
        /// <param name="allowMerge">If set to <c>true</c> allow merge the same value cells.</param>
        public PropertyConfiguration<TModel, TProperty> IsMergeEnabled(bool allowMerge)
        {
            return this;
        }
    }
}
