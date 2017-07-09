// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace Arch.FluentExcel
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents the all setting for save to and loading from excel.
    /// </summary>
    public class ExcelSetting
    {
        /// <summary>
        /// Gets or sets the comany name property of the generated excel file.
        /// </summary>
        public string Company { get; set; } = "rigofunc (yingtingxu)";

        /// <summary>
        /// Gets or sets the author property of the generated excel file.
        /// </summary>
        public string Author { get; set; } = "rigofunc (yingtingxu)";

        /// <summary>
        /// Gets or sets the subject property of the generated excel file.
        /// </summary>
        public string Subject { get; set; } = "The extensions of NPOI, which provides IEnumerable<T>; save to and load from excel.";

        /// <summary>
        /// Gets or sets a value indicating whether to use *.xlsx file extension.
        /// </summary>
        public bool UserXlsx { get; set; } = true;

        /// <summary>
        /// Gets or sets the date time formatter.
        /// </summary>
        [Obsolete("This configuration doesn't work now, please using fluent api or attribute to configure this.", true)]
        public string DateFormatter { get; set; } = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// Gets the fluent configuration entry point for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        public FluentConfiguration<TModel> For<TModel>() where TModel : class
        {
            var mc = new FluentConfiguration<TModel>();

            FluentConfigs[typeof(TModel)] = mc;

            return mc;
        }

        /// <summary>
        /// Gets the model fluent configs.
        /// </summary>
        /// <value>The model fluent configs.</value>
        internal IDictionary<Type, IFluentConfiguration> FluentConfigs { get; } = new Dictionary<Type, IFluentConfiguration>();
    }
}
