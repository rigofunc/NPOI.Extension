// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Represents the fluent configuration for the specfidied model.
    /// </summary>
    /// <typeparam name="TModel">The type of model.</typeparam>
    public class FluentConfiguration<TModel> : IFluentConfiguration where TModel : class
    {
        private Dictionary<string, PropertyConfiguration> _propertyConfigurations;
        private List<StatisticsConfiguration> _statisticsConfigurations;
        private List<FilterConfiguration> _filterConfigurations;
        private List<FreezeConfiguration> _freezeConfigurations;
        private RowDataValidator _rowDataValidator;
        private bool _skipInvalidRows;
        private bool _ignoreWhitespaceRows;

        /// <summary>
        /// Initializes a new instance of the <see cref="FluentConfiguration{TModel}"/> class.
        /// </summary>
        internal FluentConfiguration()
        {
            _propertyConfigurations = new Dictionary<string, PropertyConfiguration>();
            _statisticsConfigurations = new List<StatisticsConfiguration>();
            _filterConfigurations = new List<FilterConfiguration>();
            _freezeConfigurations = new List<FreezeConfiguration>();
        }

        /// <summary>
        /// Gets the property configurations.
        /// </summary>
        /// <value>The property configs.</value>
        public IReadOnlyDictionary<string, PropertyConfiguration> PropertyConfigurations
        {
            get
            {
                return _propertyConfigurations;
            }
        }

        /// <summary>
        /// Gets the statistics configurations.
        /// </summary>
        /// <value>The statistics config.</value>
        public IReadOnlyList<StatisticsConfiguration> StatisticsConfigurations
        {
            get
            {
                return _statisticsConfigurations.AsReadOnly();
            }
        }

        /// <summary>
        /// Gets the filter configurations.
        /// </summary>
        /// <value>The filter config.</value>
        public IReadOnlyList<FilterConfiguration> FilterConfigurations
        {
            get
            {
                return _filterConfigurations.AsReadOnly();
            }
        }

        /// <summary>
        /// Gets the freeze configurations.
        /// </summary>
        /// <value>The freeze config.</value>
        public IReadOnlyList<FreezeConfiguration> FreezeConfigurations
        {
            get
            {
                return _freezeConfigurations.AsReadOnly();
            }
        }

        /// <summary>
        /// Gets the row data validator.
        /// </summary>
        /// <value>The row data validator.</value>
        public RowDataValidator RowDataValidator { get { return _rowDataValidator; } }

        /// <summary>
        /// Gets the value indicating whether to skip the rows with validation failure while loading the excel data.
        /// </summary>
        /// <value>whether to skip the rows with validation failure</value>
        public bool SkipInvalidRows { get { return _skipInvalidRows; } }

        /// <summary>
        /// Gets the value indicating whether to ignore the rows whose cells are all blank or whitespace.
        /// </summary>
        /// <value>whether to ignore the rows whose cells are all blank or whitespace</value>
        public bool IgnoreWhitespaceRows { get { return _ignoreWhitespaceRows; } }

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TModel"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public PropertyConfiguration Property<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            var propertyInfo = GetPropertyInfo(propertyExpression);

            if (!_propertyConfigurations.TryGetValue(propertyInfo.Name, out var pc))
            {
                pc = new PropertyConfiguration();
                _propertyConfigurations[propertyInfo.Name] = pc;
            }

            return pc;
        }

        /// <summary>
        /// Gets the property configuration by the specified property info for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <param name="propertyInfo">The property information.</param>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        public PropertyConfiguration Property(PropertyInfo propertyInfo)
        {
            if (propertyInfo.DeclaringType != typeof(TModel))
            {
                throw new InvalidOperationException($"Property does not belong to {nameof(TModel)}");
            }

            if (!_propertyConfigurations.TryGetValue(propertyInfo.Name, out var pc))
            {
                pc = new PropertyConfiguration();
                _propertyConfigurations[propertyInfo.Name] = pc;
            }

            return pc;
        }

        /// <summary>
        /// Configures the ignored properties for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <param name="propertyExpressions">The a range of the property expression.</param>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        public FluentConfiguration<TModel> HasIgnoredProperties(params Expression<Func<TModel, object>>[] propertyExpressions)
        {
            foreach (var propertyExpression in propertyExpressions)
            {
                var propertyInfo = GetPropertyInfo(propertyExpression);

                if (!_propertyConfigurations.TryGetValue(propertyInfo.Name, out var pc))
                {
                    pc = new PropertyConfiguration();
                    _propertyConfigurations[propertyInfo.Name] = pc;
                }

                pc.IsIgnored(true, true);
            }

            return this;
        }

        /// <summary>
        /// Adjust the auto index value for all the has auto index configuration properties of specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        public FluentConfiguration<TModel> AdjustAutoIndex()
        {
            var index = 0;
            var properties = typeof(TModel).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);
            foreach (var property in properties)
            {
                if (!_propertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    if (_propertyConfigurations.Values.Any(c => c.Index == index))
                    {
                        // the default index had been used, so calculate a new one for it.
                        _propertyConfigurations[property.Name] = pc = new PropertyConfiguration
                        {
                            Title = property.Name,
                            AutoIndex = true,
                            Index = -1
                        };
                    }
                    else
                    {
                        // the default index not be used, 'I' will use it.
                        index++;

                        continue;
                    }
                }

                if (pc.AutoIndex && !pc.IsExportIgnored && pc.Index == -1)
                {
                    while (_propertyConfigurations.Values.Any(c => c.Index == index))
                    {
                        index++;
                    }

                    pc.HasExcelIndex(index++);
                }
            }

            return this;
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
            var statistics = new StatisticsConfiguration
            {
                Name = name,
                Formula = formula,
                Columns = columnIndexes,
            };

            _statisticsConfigurations.Add(statistics);

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
            var filter = new FilterConfiguration
            {
                FirstCol = firstColumn,
                FirstRow = firstRow,
                LastCol = lastColumn,
                LastRow = lastRow,
            };

            _filterConfigurations.Add(filter);

            return this;
        }

        /// <summary>
        /// Configures the excel freeze behaviors for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="columnSplit">The column number to split.</param>
        /// <param name="rowSplit">The row number to split.</param>
        /// <param name="leftMostColumn">The left most culomn index.</param>
        /// <param name="topMostRow">The top most row index.</param>
        public FluentConfiguration<TModel> HasFreeze(int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
        {
            var freeze = new FreezeConfiguration
            {
                ColSplit = columnSplit,
                RowSplit = rowSplit,
                LeftMostColumn = leftMostColumn,
                TopRow = topMostRow,
            };

            _freezeConfigurations.Add(freeze);

            return this;
        }

        /// <summary>
        /// Configures the row data validator which validates each row before adding it to the result list.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="rowDataValidator">The row data validator</param>
        public FluentConfiguration<TModel> HasRowDataValidator(RowDataValidator<TModel> rowDataValidator)
        {
            if (null == rowDataValidator)
            {
                _rowDataValidator = null;
                return this;
            }

            _rowDataValidator = (rowIndex, rowData) =>
            {
                var model = rowData as TModel;
                if (null == model && null != rowData) throw new ArgumentException($"the row data is not of type {typeof(TModel).Name}", nameof(rowData));

                return rowDataValidator(rowIndex, model);
            };

            return this;
        }

        /// <summary>
        /// Configure whether to skip the rows with validation failure while loading the excel data.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="shouldSkip">whether to skip</param>
        public FluentConfiguration<TModel> ShouldSkipInvalidRows(bool shouldSkip = false)
        {
            _skipInvalidRows = shouldSkip;

            return this;
        }

        /// <summary>
        /// Configure whether to ignore the rows whose cells are all blank or whitespace while loading the excel data.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="shouldIgnore">whether to ignore</param>
        public FluentConfiguration<TModel> ShouldIgnoreWhitespaceRows(bool shouldIgnore = true)
        {
            _ignoreWhitespaceRows = shouldIgnore;

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
