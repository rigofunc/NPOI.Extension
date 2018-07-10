// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;

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
        public string Subject { get; set; } = "The extensions of NPOI, which provides IEnumerable<T> has save to and load from excel functionalities.";

        /// <summary>
        /// Gets or sets a value indicating whether to use *.xlsx file extension.
        /// </summary>
        public bool UseXlsx { get; set; } = true;

        /// <summary>
        /// Gets or sets a valude indicating whether to autosize the columns. Recommmended to disable this for performance issues if the amount of data is huge.
        /// </summary>
        public bool AutoSizeColumnsEnabled { get; set; } = true;

        /// <summary>
        /// Gets or sets the title cell style applier.
        /// </summary>
        /// <value>The title cell style applier.</value>
        public Action<ICellStyle, IFont> TitleCellStyleApplier { get; set; } = DefaultTitleCellStyleApplier;

        /// <summary>
        /// Gets the fluent configuration entry point for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        /// <param name="refreshCache"><c>True</c> if to refresh cache, ortherwise, <c>false</c>.</param>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        public FluentConfiguration<TModel> For<TModel>(bool refreshCache = false) where TModel : class
        {
            var type = typeof(TModel);
            if (!FluentConfigs.TryGetValue(type, out var mc) || refreshCache)
            {
                mc = new FluentConfiguration<TModel>();

                FluentConfigs[type] = mc;
            }

            return mc as FluentConfiguration<TModel>;
        }

        /// <summary>
        /// Gets the model fluent configs.
        /// </summary>
        /// <value>The model fluent configs.</value>
        internal IDictionary<Type, IFluentConfiguration> FluentConfigs { get; } = new Dictionary<Type, IFluentConfiguration>();

        internal static void DefaultTitleCellStyleApplier(ICellStyle cellStyle, IFont font)
        {
            cellStyle.Alignment = HorizontalAlignment.Center;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            font.Boldweight = (short)FontBoldWeight.Bold;
            cellStyle.SetFont(font);
        }
    }
}
