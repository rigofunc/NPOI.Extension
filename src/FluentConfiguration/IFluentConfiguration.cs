// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System.Collections.Generic;
    using System.Reflection;

    /// <summary>
    /// Provides the interfaces for the fluent configuration.
    /// </summary>
    internal interface IFluentConfiguration
    {
        /// <summary>
        /// Gets the property configs.
        /// </summary>
        /// <value>The property configs.</value>
        IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigs { get; }

        /// <summary>
        /// Gets the statistics configs.
        /// </summary>
        /// <value>The statistics config.</value>
        IList<StatisticsConfig> StatisticsConfigs { get; }

        /// <summary>
        /// Gets the filter configs.
        /// </summary>
        /// <value>The filter config.</value>
        IList<FilterConfig> FilterConfigs { get; }

        /// <summary>
        /// Gets the freeze configs.
        /// </summary>
        /// <value>The freeze config.</value>
        IList<FreezeConfig> FreezeConfigs { get; }
    }
}
