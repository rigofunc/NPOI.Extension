// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace NPOI.Extension
{
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
        public string DateFormatter { get; set; } = "yyyy-MM-dd HH:mm:ss";
    }
}
