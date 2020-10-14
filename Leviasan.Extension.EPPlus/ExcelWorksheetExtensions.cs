using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using OfficeOpenXml.Data.Annotations;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace OfficeOpenXml
{
    /// <summary>
    /// The excel worksheet extension methods.
    /// </summary>
    public static class ExcelWorksheetExtensions
    {
        /// <summary>
        /// Loads a dictionary value into a worksheet into a range that contains a dictionary key.
        /// </summary>
        /// <param name="worksheet">The excel worksheet.</param>
        /// <param name="dictionary">The dictionary with value.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        /// <exception cref="ArgumentNullException">Worksheet or dictionary is null.</exception>
        public static ExcelWorksheet LoadFromDictionary(this ExcelWorksheet worksheet, IDictionary<string, object> dictionary, IFormatProvider provider = default)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (dictionary == null)
                throw new ArgumentNullException(nameof(dictionary));

            foreach (var element in dictionary)
            {
                var key = FormatKey(element.Key);
                if (worksheet.Cells.FirstOrDefault(x => x.Text.Contains(key, StringComparison.OrdinalIgnoreCase)) is ExcelRangeBase excelRange)
                    SetExcelRangeValue(excelRange, key, element.Value, true, provider);
            }
            return worksheet;
        }
        /// <summary>
        /// Loads a object into a worksheet into a range that contains a [property name].
        /// </summary>
        /// <param name="worksheet">The excel worksheet.</param>
        /// <param name="value">The object.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        /// <exception cref="ArgumentNullException">Worksheet or dictionary is null.</exception>
        public static ExcelWorksheet LoadFromDictionary(this ExcelWorksheet worksheet, object value, IFormatProvider provider = default)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            var properties = value.GetType().GetProperties();
            foreach (var property in properties)
            {
                var key = FormatKey(property.Name);
                if (worksheet.Cells.FirstOrDefault(x => x.Text.Contains(key, StringComparison.OrdinalIgnoreCase)) is ExcelRangeBase excelRange)
                    SetExcelRangeValue(excelRange, property, property.GetValue(value), provider);
            }
            return worksheet;
        }
        /// <summary>
        /// Loads a collection into a worksheet starting from the top left row of the range that contains keyword.
        /// </summary>
        /// <typeparam name="T">The type of data.</typeparam>
        /// <param name="worksheet">The excel worksheet.</param>
        /// <param name="keyword">The keyword to search in the worksheet to determine the starting position.</param>
        /// <param name="collection">The data.</param>
        /// <param name="printHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="tableStyles">Will create a table with this style.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        /// <exception cref="ArgumentException">Keyword is null, empty, or consists only of white-space characters.</exception>
        /// <exception cref="ArgumentNullException">Worksheet or collection is null.</exception>
        /// <exception cref="InvalidOperationException">Keyword is not found in excel worksheet.</exception>
        public static ExcelWorksheet LoadFromCollection<T>(this ExcelWorksheet worksheet, string keyword, IEnumerable<T> collection, bool printHeaders = false, TableStyles tableStyles = TableStyles.None, IFormatProvider provider = default) where T : class
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (string.IsNullOrWhiteSpace(keyword))
                throw new ArgumentException(Properties.Resources.StringIsMissing, nameof(collection));
            if (collection == null)
                throw new ArgumentNullException(nameof(collection));

            if (worksheet.Cells.FirstOrDefault(x => x.Text.Contains(FormatKey(keyword), StringComparison.OrdinalIgnoreCase)) is ExcelRangeBase range)
                return LoadFromCollection(worksheet, range.Start.Row, range.Start.Column, collection, printHeaders, tableStyles, provider);

            throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, Properties.Resources.KeywordNotFoundException, worksheet.Name, keyword));
        }
        /// <summary>
        /// Load a collection into a worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The type of data.</typeparam>
        /// <param name="worksheet">The excel worksheet.</param>
        /// <param name="startRow">The start number row.</param>
        /// <param name="startColumn">The start number column.</param>
        /// <param name="collection">The data.</param>
        /// <param name="printHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="tableStyles">Will create a table with this style.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        /// <exception cref="ArgumentNullException">Worksheet or collection is null.</exception>
        public static ExcelWorksheet LoadFromCollection<T>(this ExcelWorksheet worksheet, int startRow, int startColumn, IEnumerable<T> collection, bool printHeaders = false, TableStyles tableStyles = TableStyles.None, IFormatProvider provider = default) where T : class
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (collection == null)
                throw new ArgumentNullException(nameof(collection));

            // Get property info
            var properties = typeof(T).GetProperties();

            // Get position
            var start = worksheet.Cells[startRow, startColumn];
            var next = worksheet.Cells[start.Address];

            // Print headers
            if (printHeaders)
            {
                foreach (var property in properties)
                {
                    // Check print hyperlink if it allow insert 
                    if (Attribute.GetCustomAttribute(property, typeof(HyperlinkAttribute)) is HyperlinkAttribute hyperlink && !hyperlink.AllowInsertValue)
                        continue;

                    // Get display value
                    var value = Attribute.GetCustomAttribute(property, typeof(DisplayNameAttribute)) is DisplayNameAttribute display
                        ? display.DisplayName
                        : property.Name;

                    // Set value
                    SetExcelRangeValue(next, property, value, provider);
                    next.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // Move next column
                    next = worksheet.Cells[next.Start.Row, next.Start.Column + 1];
                }
                // Move next row
                next = worksheet.Cells[next.Start.Row + 1, start.Start.Column];
            }

            // Print collection
            var updateNext = false;
            foreach (var obj in collection)
            {
                // Move next row
                if (updateNext)
                {
                    worksheet.InsertRow(next.Start.Row + 1, 1, 1);
                    next = worksheet.Cells[next.Start.Row + 1, start.Start.Column];
                }
                foreach (var property in properties)
                {
                    // The flag which allows the insertion of a value
                    var allowInsertValue = true;
                    // Get property value
                    var value = property.GetValue(obj);
                    // Set hyperlink and print it if allowed
                    if (Attribute.GetCustomAttribute(property, typeof(HyperlinkAttribute)) is HyperlinkAttribute hyperlink)
                    {
                        // Condition: url is not null or empty.
                        if (!string.IsNullOrWhiteSpace(value as string))
                        {
                            // Get index hyperlink property
                            var propertyName = hyperlink.Property ?? property.Name;
                            var index = properties.ToList().FindIndex(p => p.Name == propertyName);
                            // Get range where property printed in a row
                            var range = worksheet.Cells[next.Start.Row, start.Start.Column + index];
                            // Set value
                            SetExcelRangeValue(range, property, value, provider);
                        }
                        // Definition grants
                        allowInsertValue = hyperlink.AllowInsertValue;
                    }
                    if (allowInsertValue)
                    {
                        // Grouping data
                        if (Attribute.GetCustomAttribute(property, typeof(GroupingAttribute)) is GroupingAttribute grouping)
                        {
                            // Condition: if text value in the previous row in the current column is equaled current cell text value
                            if (worksheet.Cells[next.Start.Row - 1, next.Start.Column].Text.Equals(value as string, StringComparison.Ordinal))
                            {
                                worksheet.Row(next.Start.Row).OutlineLevel = grouping.OutlineLevel;
                                worksheet.Row(next.Start.Row).Collapsed = grouping.Collapsed;
                            }
                        }
                        // Set value
                        SetExcelRangeValue(next, property, value, provider);
                        // Move next column
                        next = worksheet.Cells[next.Start.Row, next.Start.Column + 1];
                    }
                }
                updateNext = true;
            }
            worksheet.InsertRow(next.Start.Row + 1, 1, 1);

            // Set table style
            var diff = printHeaders ? 0 : 1;
            var tableRange = worksheet.Cells[start.Start.Row - diff, start.Start.Column, next.Start.Row, next.Start.Column - 1];
            var table = worksheet.Tables.Add(tableRange, $"{typeof(T).Name}_{DateTime.UtcNow.Ticks}");
            table.TableStyle = tableStyles;
            // Set excel cell border style
            tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;

            return worksheet;
        }

        /// <summary>
        /// Gets string that represent the key.
        /// </summary>
        /// <param name="value">The key value.</param>
        private static string FormatKey(string value)
        {
            return $"[{value}]";
        }
        /// <summary>
        /// Sets the value in the specified excel range.
        /// </summary>
        /// <param name="range">A range of cells.</param>
        /// <param name="property">Property metadata.</param>
        /// <param name="value">The value.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        private static void SetExcelRangeValue(ExcelRangeBase range, PropertyInfo property, object value, IFormatProvider provider = default)
        {
            var styleDefinition = true;
            if (Attribute.GetCustomAttribute(property, typeof(NumberformatAttribute)) is NumberformatAttribute numberformatAttribute)
            {
                styleDefinition = false;
                range.Style.Numberformat.Format = numberformatAttribute.Format;
            }
            if (Attribute.GetCustomAttribute(property, typeof(RangeStyleAttribute)) is RangeStyleAttribute rangeStyleAttribute)
            {
                range.Style.Font.Size = rangeStyleAttribute.FontSize;
                range.Style.WrapText = rangeStyleAttribute.WrapText;
                range.Style.VerticalAlignment = rangeStyleAttribute.VerticalAlignment;
                range.Style.HorizontalAlignment = rangeStyleAttribute.HorizontalAlignment;
            }
            if (Attribute.GetCustomAttribute(property, typeof(DescriptionAttribute)) is DescriptionAttribute description)
            {
                var comment = description.Description;
                if (!string.IsNullOrWhiteSpace(comment))
                    range.AddComment(comment, nameof(OfficeOpenXml));
            }
            if (Attribute.GetCustomAttribute(property, typeof(HyperlinkAttribute)) is HyperlinkAttribute hyperlink)
            {
                // Set style
                range.StyleName = hyperlink.StyleName;
                // Set hyperlink value
                range.Hyperlink = new Uri(value as string);
                if (!hyperlink.AllowInsertValue)
                    return;
            }
            SetExcelRangeValue(range, FormatKey(property.Name), value, styleDefinition, provider);
        }
        /// <summary>
        /// Sets the value in the specified excel range.
        /// </summary>
        /// <param name="range">A range of cells.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <param name="styleDefinition">If needs to definition the style set true, overwise, false.</param>
        /// <param name="provider"><see cref="CultureInfo"/> provider.</param>
        private static void SetExcelRangeValue(ExcelRangeBase range, string key, object value, bool styleDefinition = true, IFormatProvider provider = default)
        {
            // Definition type
            var type = value?.GetType();
            // Sets style if it needs
            if (styleDefinition && value != null)
            {
                // Definition style based on type
                if (type == typeof(DateTime))
                    range.Style.Numberformat.Format = CultureInfo.InvariantCulture.DateTimeFormat.FullDateTimePattern;
                else if (type == typeof(float) || type == typeof(double) || type == typeof(decimal))
                    range.Style.Numberformat.Format = "0.00";
                else if (type == typeof(char) || type == typeof(string))
                    range.Style.Numberformat.Format = "@";
                else if (type.IsPrimitive)
                    range.Style.Numberformat.Format = "0";
            }
            // Define replacement
            var replacement = value;
            if (type != null)
            {
                replacement = type.Equals(typeof(DateTime))
                    ? Convert.ToDateTime(value, provider).ToString(range.Style.Numberformat.Format, provider)
                    : Convert.ChangeType(value, type, provider);
            }
            // Set value or replace by pattern
            range.Value = range.Text.Contains(key, StringComparison.Ordinal)
                ? range.Text.Replace(key, replacement?.ToString(), StringComparison.Ordinal)
                : value;
        }
    }
}
