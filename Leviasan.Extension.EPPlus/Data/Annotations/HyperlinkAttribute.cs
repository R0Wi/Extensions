using System;

namespace OfficeOpenXml.Data.Annotations
{
    /// <summary>
    /// Represents the hyperlink property in table data.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class HyperlinkAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of <see cref="HyperlinkAttribute"/> class with current property. The property <see cref="AllowInsertValue"/> by default is true.
        /// </summary>
        public HyperlinkAttribute()
        {
            AllowInsertValue = true;
        }
        /// <summary>
        /// Initializes a new instance of <see cref="HyperlinkAttribute"/> class with specified property name that value is set how hyperlink. 
        /// Also sets the value in next column if <see cref="AllowInsertValue"/> is true. 
        /// </summary>
        /// <param name="property">The name of the property where need set the value how hyperlink.</param>
        /// <param name="allowInsertValue">Is sets the value in the next column.</param>
        /// <exception cref="ArgumentException">Property is null, empty, or consists only of white-space characters.</exception>
        public HyperlinkAttribute(string property, bool allowInsertValue)
        {
            if (string.IsNullOrWhiteSpace(property))
                throw new ArgumentException(Properties.Resources.StringIsMissing, nameof(property));

            AllowInsertValue = allowInsertValue;
            Property = property;
        }

        /// <summary>
        /// Is sets the value in the next column.
        /// </summary>
        public bool AllowInsertValue { get; }
        /// <summary>
        /// The name of the property where need set the value how hyperlink.
        /// </summary>
        public string Property { get; }
        /// <summary>
        /// The hyperlink style name.
        /// </summary>
        public string StyleName { get; set; } = "Hyperlink";
    }
}
