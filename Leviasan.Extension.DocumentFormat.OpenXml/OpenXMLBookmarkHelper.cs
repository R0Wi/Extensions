using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml
{
    /// <summary>
    /// Represents the service that working with OpenXML bookmarks.
    /// </summary>
    public static class OpenXmlBookmarkHelper
    {
        /// <summary>
        /// Gets the all document bookmarks values.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="includeHiddenBookmarks"></param>
        /// <param name="provider"></param>
        /// <exception cref="ArgumentNullException">The <see cref="WordprocessingDocument"/> is null.</exception>
        public static IDictionary<string, string> GetDocumentBookmarkValues(WordprocessingDocument document, bool includeHiddenBookmarks = false, IFormatProvider provider = default)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var bookmarks = new Dictionary<string, string>();
            foreach (var bookmark in GetAllBookmarks(document))
            {
                if (includeHiddenBookmarks || !IsHiddenBookmark(bookmark.Name))
                    bookmarks[bookmark.Name] = GetValue(bookmark, provider);
            }
            return bookmarks;
        }
        /// <summary>
        /// Sets the values into the document bookmarks.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="bookmarkValues"></param>
        /// <param name="provider"></param>
        /// <exception cref="ArgumentNullException">The document or dictionary is null.</exception>
        public static void SetDocumentBookmarkValues(WordprocessingDocument document, IDictionary<string, string> bookmarkValues, IFormatProvider provider = default)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (bookmarkValues == null)
                throw new ArgumentNullException(nameof(bookmarkValues));

            foreach (var bookmark in GetAllBookmarks(document))
                SetBookmarkValue(bookmark, bookmarkValues, provider);
        }
        /// <summary>
        /// Gets the value from the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="provider"></param>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        public static string GetValue(BookmarkStart bookmark, IFormatProvider provider = default)
        {
            if (bookmark == null)
                throw new ArgumentNullException(nameof(bookmark));

            var formFieldData = bookmark.Parent
                .Descendants<FormFieldData>()
                .FirstOrDefault(x => x.ChildElements
                    .Any(el => el is FormFieldName formFieldName && string.Equals(formFieldName.Val.Value, bookmark.Name, StringComparison.OrdinalIgnoreCase)));

            if (formFieldData != null)
            {
                var checkbox = formFieldData.GetFirstDescendant<DefaultCheckBoxFormFieldState>();
                if (checkbox != null)
                    return Convert.ToString(checkbox.Val.Value, provider);
            }

            var separate = bookmark.Parent.Descendants<FieldChar>().FirstOrDefault(x => x.FieldCharType == FieldCharValues.Separate && x.IsAfter(bookmark));
            var nextSibling = separate.Parent;
            while (nextSibling != null)
            {
                if (nextSibling.IsEndBookmark(bookmark))
                    break;

                var next = nextSibling.NextSibling();
                if (nextSibling.IsAfter(separate))
                {
                    if (nextSibling.ChildElements.Any(x => x.GetType().Equals(typeof(Text))))
                    {
                        var text = nextSibling.GetFirstChild<Text>();
                        return text.Text;
                    }
                }
                nextSibling = next;
            }

            return null;
        }
        /// <summary>
        /// Sets the value into the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="value"></param>
        /// <param name="provider"></param>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        public static void SetValue(BookmarkStart bookmark, string value, IFormatProvider provider = default)
        {
            if (bookmark == null)
                throw new ArgumentNullException(nameof(bookmark));

            var formFieldData = bookmark.Parent
                .Descendants<FormFieldData>()
                .FirstOrDefault(x => x.ChildElements
                    .Any(el => el is FormFieldName formFieldName && string.Equals(formFieldName.Val.Value, bookmark.Name, StringComparison.OrdinalIgnoreCase)));

            if (formFieldData != null)
            {
                // TextInput
                var textInput = formFieldData.GetFirstDescendant<TextInput>();
                if (textInput != null)
                {
                    // Get default value from bookmark if the settable value is null
                    if (value == null)
                    {
                        var defaultValue = formFieldData.GetFirstDescendant<DefaultTextBoxFormFieldString>();
                        if (defaultValue != null && defaultValue.Val.HasValue)
                            value = defaultValue.Val.Value;
                    }

                    // Formating datetime string
                    var format = formFieldData.GetFirstDescendant<Format>();
                    var formFieldType = formFieldData.GetFirstDescendant<TextBoxFormFieldType>();
                    if (formFieldType != null && format != null && formFieldType.Val.Value == TextBoxFormFieldValues.Date && !string.IsNullOrWhiteSpace(value))
                    {
                        // Try convert with specified format provider
                        if (DateTime.TryParse(value, provider, DateTimeStyles.None, out var datetimeCurrent))
                        {
                            value = format.Val.HasValue
                                ? datetimeCurrent.ToString(format.Val.Value, provider)
                                : value;
                        }
                        else
                        {
                            throw new FormatException("A string that represents the datetime does not contain a valid string representation of a date and time or using an invalid format provider.");
                        }
                    }

                    // Enforce max length.
                    var maxLength = formFieldData.GetFirstDescendant<MaxLength>();
                    if (maxLength != null && maxLength.Val.HasValue)
                        value = value.Substring(0, maxLength.Val.Value);

                    // Removes other empty run elements after separate
                    var separate = bookmark.Parent.Descendants<FieldChar>().FirstOrDefault(x => x.FieldCharType == FieldCharValues.Separate && x.IsAfter(bookmark));
                    var nextSibling = separate.Parent;
                    RunProperties runProperties = null;
                    while (nextSibling != null)
                    {
                        if (nextSibling.IsEndBookmark(bookmark))
                            break;

                        var next = nextSibling.NextSibling();
                        if (nextSibling.IsAfter(separate))
                        {
                            if (nextSibling.ChildElements.Any(x => x.GetType().Equals(typeof(Text))))
                            {
                                if (runProperties == null)
                                    runProperties = nextSibling.GetFirstChild<RunProperties>().Clone() as RunProperties;

                                nextSibling.Remove();
                            }
                        }
                        nextSibling = next;
                    }
                    // Set value
                    var end = bookmark.Parent.Descendants<FieldChar>().FirstOrDefault(x => x.FieldCharType == FieldCharValues.End && x.IsAfter(bookmark)).Parent;
                    var text = new Text(value);
                    var run = new Run(runProperties, text);
                    bookmark.Parent.InsertBefore(run, end);
                    return;
                }
                // Checkbox
                var checkbox = formFieldData.GetFirstDescendant<DefaultCheckBoxFormFieldState>();
                if (checkbox != null)
                {
                    checkbox.Val = new OnOffValue(bool.TryParse(value, out var result)
                        ? result
                        : Convert.ToBoolean(Convert.ToInt32(value, provider), provider));
                    return;
                }
            }
        }

        /// <summary>
        /// Gets all bookmarks in document.
        /// </summary>
        /// <param name="document"></param>
        private static IEnumerable<BookmarkStart> GetAllBookmarks(WordprocessingDocument document)
        {
            return document.MainDocumentPart.RootElement.Descendants<BookmarkStart>();
        }
        /// <summary>
        /// Is hidden bookmark The hidden bookmark name starts with '_'.
        /// </summary>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Return true if the bookmark is hidden, otherwise false.</returns>
        private static bool IsHiddenBookmark(string bookmarkName)
        {
            return bookmarkName?.StartsWith("_", StringComparison.OrdinalIgnoreCase) ?? false;
        }
        /// <summary>
        /// Sets the value to bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="bookmarkValues"></param>
        /// <param name="provider"></param>
        private static void SetBookmarkValue(BookmarkStart bookmark, IDictionary<string, string> bookmarkValues, IFormatProvider provider = default)
        {
            if (bookmarkValues.TryGetValue(bookmark.Name, out var value))
                SetValue(bookmark, value, provider);
        }
    }
}
