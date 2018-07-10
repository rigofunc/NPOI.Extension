// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System.Collections.Generic;

    /// <summary>
    /// Row data validator delegate, validate current row before adding it to the result list.
    /// </summary>
    /// <param name="rowIndex">Index of current row in excel</param>
    /// <param name="rowData">Model data of current row</param>
    /// <returns>Whether the row data passes validation</returns>
    public delegate bool RowDataValidatorDelegate(int rowIndex, object rowData);

    /// <summary>
    /// Provides the interfaces for the fluent configuration.
    /// </summary>
    public interface IFluentConfiguration
    {
        /// <summary>
        /// Gets the property configurations.
        /// </summary>
        /// <value>The property configs.</value>
        IReadOnlyDictionary<string, PropertyConfiguration> PropertyConfigurations { get; }

        /// <summary>
        /// Gets the statistics configurations.
        /// </summary>
        /// <value>The statistics config.</value>
        IReadOnlyList<StatisticsConfiguration> StatisticsConfigurations { get; }

        /// <summary>
        /// Gets the filter configurations.
        /// </summary>
        /// <value>The filter config.</value>
        IReadOnlyList<FilterConfiguration> FilterConfigurations { get; }

        /// <summary>
        /// Gets the freeze configurations.
        /// </summary>
        /// <value>The freeze config.</value>
        IReadOnlyList<FreezeConfiguration> FreezeConfigurations { get; }

        /// <summary>
        /// Gets the row data validator.
        /// </summary>
        /// <value>The row data validator.</value>
        RowDataValidatorDelegate RowDataValidator { get; }

        /// <summary>
        /// Gets the value indicating whether to skip the rows with validation failure while loading the excel data.
        /// </summary>
        bool SkipInvalidRows { get; }
    }
}
