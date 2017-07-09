// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Represents the fluent configuration for the specfidied model.
    /// </summary>
    /// <typeparam name="TModel">The type of model.</typeparam>
    public class FluentConfiguration<TModel> : IFluentConfiguration where TModel : class
    {
        private IDictionary<PropertyInfo, PropertyConfiguration> _propertyConfigs;
        private IList<StatisticsConfig> _statisticsConfigs;
        private IList<FilterConfig> _filterConfigs;
        private IList<FreezeConfig> _freezeConfigs;

        /// <summary>
        /// Initializes a new instance of the <see cref="FluentConfiguration{TModel}"/> class.
        /// </summary>
        public FluentConfiguration()
        {
            _propertyConfigs = new Dictionary<PropertyInfo, PropertyConfiguration>();
            _statisticsConfigs = new List<StatisticsConfig>();
            _filterConfigs = new List<FilterConfig>();
            _freezeConfigs = new List<FreezeConfig>();
        }

        /// <summary>
        /// Gets the property configs.
        /// </summary>
        /// <value>The property configs.</value>
        IDictionary<PropertyInfo, PropertyConfiguration> IFluentConfiguration.PropertyConfigs
        {
            get
            {
                return _propertyConfigs;
            }
        }

        /// <summary>
        /// Gets the statistics configs.
        /// </summary>
        /// <value>The statistics config.</value>
        IList<StatisticsConfig> IFluentConfiguration.StatisticsConfigs
        {
            get
            {
                return _statisticsConfigs;
            }
        }

        /// <summary>
        /// Gets the filter configs.
        /// </summary>
        /// <value>The filter config.</value>
        IList<FilterConfig> IFluentConfiguration.FilterConfigs
        {
            get
            {
                return _filterConfigs;
            }
        }

        /// <summary>
        /// Gets the freeze configs.
        /// </summary>
        /// <value>The freeze config.</value>
        IList<FreezeConfig> IFluentConfiguration.FreezeConfigs
        {
            get
            {
                return _freezeConfigs;
            }
        }

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TModel"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public PropertyConfiguration Property<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            var pc = new PropertyConfiguration();

            var propertyInfo = GetPropertyInfo(propertyExpression);

            _propertyConfigs[propertyInfo] = pc;

            return pc;
        }

        /// <summary>
        /// Configures the statistics for the specified <typeparamref name="TModel"/>. Only for vertical, not for horizontal statistics.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="name">The statistics name. (e.g. Total). In current version, the default name location is (last row, first cell)</param>
        /// <param name="formula">The cell formula, such as SUM, AVERAGE and so on, which applyable for vertical statistics..</param>
        /// <param name="columnIndexes">The column indexes for statistics. if <paramref name="formula"/>is SUM, and <paramref name="columnIndexes"/> is [1,3], 
        /// for example, the column No. 1 and 3 will be SUM for first row to last row.</param>
        public FluentConfiguration<TModel> HasStatistics(string name, string formula, params int[] columnIndexes)
        {
            var statistics = new StatisticsConfig
            {
                Name = name,
                Formula = formula,
                Columns = columnIndexes,
            };

            _statisticsConfigs.Add(statistics);

            return this;
        }

        /// <summary>
        /// Configures the excel filter behaviors for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="firstColumn">The first column index.</param>
        /// <param name="lastColumn">The last column index.</param>
        /// <param name="firstRow">The first row index.</param>
        /// <param name="lastRow">The last row index. If is null, the value is dynamic calculate by code.</param>
        public FluentConfiguration<TModel> HasFilter(int firstColumn, int lastColumn, int firstRow, int? lastRow = null)
        {
            var filter = new FilterConfig
            {
                FirstCol = firstColumn,
                FirstRow = firstRow,
                LastCol = lastColumn,
                LastRow = lastRow,
            };

            _filterConfigs.Add(filter);

            return this;
        }

        /// <summary>
        /// Configures the excel freeze behaviors for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="columnSplit">The column number to split.</param>
        /// <param name="rowSplit">The row number to split.param>
        /// <param name="leftMostColumn">The left most culomn index.</param>
        /// <param name="topMostRow">The top most row index.</param>
        public FluentConfiguration<TModel> HasFreeze(int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
        {
            var freeze = new FreezeConfig
            {
                ColSplit = columnSplit,
                RowSplit = rowSplit,
                LeftMostColumn = leftMostColumn,
                TopRow = topMostRow,
            };

            _freezeConfigs.Add(freeze);

            return this;
        }

        private PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            if (propertyExpression.NodeType != ExpressionType.Lambda)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            var lambda = (LambdaExpression)propertyExpression;

            var memberExpression = ExtractMemberExpression(lambda.Body);
            if (memberExpression == null)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            if (memberExpression.Member.DeclaringType == null)
            {
                throw new InvalidOperationException("Property does not have declaring type");
            }

            return memberExpression.Member.DeclaringType.GetProperty(memberExpression.Member.Name);
        }

        private MemberExpression ExtractMemberExpression(Expression expression)
        {
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                return ((MemberExpression)expression);
            }

            if (expression.NodeType == ExpressionType.Convert)
            {
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            }

            return null;
        }
    }
}