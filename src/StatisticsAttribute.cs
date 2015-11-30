// Copyright (c) RigoFunc (xuyingting). All rights reserved.

namespace NPOI.Extension {
    using System;

    /// <summary>
    /// Represents a custom attribute for some simple statistics.
    /// </summary>
    /// <remarks>
    /// Only for vertical, not for horizontal statistics. and in current version, 
    /// doesn't allow apply multiple <see cref="StatisticsAttribute"/> to one class.
    /// </remarks>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class StatisticsAttribute : Attribute {
        /// <summary>
        /// Gets or sets the statistics name. (e.g. Total)
        /// </summary>
        /// <remarks>
        /// In current version, the default name location is (last row, first cell)
        /// </remarks>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the cell formula, such as SUM, AVERAGE and so on, which applyable for vertical statistics.
        /// </summary>
        public string Formula { get; set; }

        /// <summary>
        /// Gets or sets the column indexes for statistics. if <see cref="Formula"/> is SUM, 
        /// and <see cref="Columns"/> is [1,3], for example, the column No. 1 and 3 will be
        /// SUM for first row to last row.
        /// </summary>
        public int[] Columns { get; set; }
    }
}
