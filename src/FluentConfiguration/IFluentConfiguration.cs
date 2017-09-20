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
        /// Gets the property configs.
        /// </summary>
        /// <value>The property configs.</value>
        IReadOnlyDictionary<string, PropertyConfiguration> PropertyConfigs { get; }

        /// <summary>
        /// Gets the statistics configs.
        /// </summary>
        /// <value>The statistics config.</value>
        IReadOnlyList<StatisticsConfiguration> StatisticsConfigs { get; }

        /// <summary>
        /// Gets the filter configs.
        /// </summary>
        /// <value>The filter config.</value>
        IReadOnlyList<FilterConfiguration> FilterConfigs { get; }

        /// <summary>
        /// Gets the freeze configs.
        /// </summary>
        /// <value>The freeze config.</value>
        IReadOnlyList<FreezeConfiguration> FreezeConfigs { get; }
    }
}
