using System;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Data.Annotations
{
    /// <summary>
    /// Represents the style of range.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class RangeStyleAttribute : Attribute
    {
        /// <summary>
        /// The size of the font. By gefault: 11.
        /// </summary>
        public int FontSize { get; set; } = 11;
        /// <summary>
        /// Wrap the text.
        /// </summary>
        public bool WrapText { get; set; }
        /// <summary>
        /// The vertical alignment in the cell.
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; }
        /// <summary>
        /// The horizontal alignment in the cell.
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }
    }
}
