using System;

namespace OfficeOpenXml.Data.Annotations
{
    /// <summary>
    /// Represents a data format in which the value will be displayed.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class NumberformatAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of <see cref="NumberformatAttribute"/> class with specified format.
        /// </summary>
        /// <param name="format">The data format in which the value will be displayed.</param>
        /// <exception cref="ArgumentException">Property is null, empty, or consists only of white-space characters.</exception>
        public NumberformatAttribute(string format)
        {
            if (string.IsNullOrWhiteSpace(format))
                throw new ArgumentException(Properties.Resources.StringIsMissing, nameof(format));

            Format = format;
        }

        /// <summary>
        /// The data format in which the value will be displayed.
        /// </summary>
        public string Format { get; }
    }
}
