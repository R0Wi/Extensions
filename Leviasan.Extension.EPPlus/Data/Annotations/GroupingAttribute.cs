using System;

namespace OfficeOpenXml.Data.Annotations
{
    /// <summary>
    /// Represents the grouping property in table data.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class GroupingAttribute : Attribute
    {
        /// <summary>
        /// Is row collapsed. By default is true.
        /// </summary>
        public bool Collapsed { get; set; } = true;
        /// <summary>
        /// The nesting level. By default is 1.
        /// </summary>
        public int OutlineLevel { get; set; } = 1;
    }
}
