// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System.Collections.Generic;

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
    }
}
