using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WebAPI.Helpers
{
    /// <summary>
    /// Represents the service that worked with OpenXML bookmarks.
    /// </summary>
    /// <remarks>
    /// Based on: https://gist.github.com/pgrm/5034752
    /// </remarks>
    public class OpenXMLBookmarkHelper
    {
        /// <summary>
        /// Gets the all document bookmarks values.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="includeHiddenBookmarks"></param>
        /// <exception cref="ArgumentNullException">The <see cref="WordprocessingDocument"/> is null.</exception>
        public IDictionary<string, string> GetDocumentBookmarkValues(WordprocessingDocument document, bool includeHiddenBookmarks = false)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var bookmarks = new Dictionary<string, string>();
            foreach (var bookmark in GetAllBookmarks(document))
            {
                if (includeHiddenBookmarks || !IsHiddenBookmark(bookmark.Name))
                    bookmarks[bookmark.Name] = GetText(bookmark);
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
        public void SetDocumentBookmarkValues(WordprocessingDocument document, IDictionary<string, string> bookmarkValues, IFormatProvider provider = default)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (bookmarkValues == null)
                throw new ArgumentNullException(nameof(bookmarkValues));

            foreach (var bookmark in GetAllBookmarks(document))
                SetBookmarkValue(bookmark, bookmarkValues, provider);
        }
        /// <summary>
        /// Sets the text into the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="value"></param>
        /// <param name="provider"></param>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        public void SetText(BookmarkStart bookmark, string value, IFormatProvider provider = default)
        {
            if (bookmark == null)
                throw new ArgumentNullException(nameof(bookmark));

            var text = FindBookmarkText(bookmark);
            if (text != null)
            {
                text.Text = value;
                RemoveOtherTexts(bookmark, text);
            }
            else
            {
                var checkBox = FindBookmarkCheckBox(bookmark);
                if (checkBox != null)
                {
                    var state = checkBox.GetFirstDescendant<DefaultCheckBoxFormFieldState>();
                    state.Val = new OnOffValue(bool.TryParse(value, out var result)
                        ? result
                        : Convert.ToBoolean(Convert.ToInt32(value, provider), provider));
                }
                else
                {
                    InsertBookmarkText(bookmark, value, provider);
                }
            }
        }
        /// <summary>
        /// Gets the text from the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        public string GetText(BookmarkStart bookmark)
        {
            if (bookmark == null)
                throw new ArgumentNullException(nameof(bookmark));

            var text = FindBookmarkText(bookmark);

            if (text != null)
                return text.Text;
            else
                return string.Empty;
        }

        /// <summary>
        /// Gets all bookmarks in document.
        /// </summary>
        /// <param name="document"></param>
        private IEnumerable<BookmarkStart> GetAllBookmarks(WordprocessingDocument document)
        {
            return document.MainDocumentPart.RootElement.Descendants<BookmarkStart>();
        }
        /// <summary>
        /// Is hidden bookmark The hidden bookmark name starts with '_'.
        /// </summary>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Return true if the bookmark is hidden, otherwise false.</returns>
        private bool IsHiddenBookmark(string bookmarkName)
        {
            return bookmarkName?.StartsWith("_", StringComparison.OrdinalIgnoreCase) ?? false;
        }
        /// <summary>
        /// Sets the value to bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="bookmarkValues"></param>
        /// <param name="provider"></param>
        private void SetBookmarkValue(BookmarkStart bookmark, IDictionary<string, string> bookmarkValues, IFormatProvider provider = default)
        {
            if (bookmarkValues.TryGetValue(bookmark.Name, out var value))
                SetText(bookmark, value, provider);
        }
        /// <summary>
        /// Finds the first OpenXML text element in the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        private Text FindBookmarkText(BookmarkStart bookmark)
        {
            if (bookmark.ColumnFirst != null)
            {
                return FindTextInColumn(bookmark);
            }
            else
            {
                var run = bookmark.NextSibling<Run>();
                if (run != null)
                {
                    return run.GetFirstChild<Text>();
                }
                else
                {
                    Text text = null;
                    var nextSibling = bookmark.NextSibling();
                    while (text == null && nextSibling != null)
                    {
                        if (nextSibling.IsEndBookmark(bookmark))
                            return null;

                        text = nextSibling.GetFirstDescendant<Text>();
                        nextSibling = nextSibling.NextSibling();
                    }
                    return text;
                }
            }
        }
        /// <summary>
        /// Finds the first OpenXML checkbox element in the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        private CheckBox FindBookmarkCheckBox(BookmarkStart bookmark)
        {
            var run = bookmark.NextSibling<Run>();
            if (run != null)
            {
                var formFieldData = bookmark.Parent.Descendants<FormFieldData>().First(x => x.ChildElements.Any(el => el is FormFieldName formFieldName && formFieldName.Val.Value == bookmark.Name));
                return formFieldData.GetFirstDescendant<CheckBox>();
            }
            return null;
        }
        /// <summary>
        /// Gets the first OpenXML text element in last cell in the bookmark column.
        /// </summary>
        /// <param name="bookmark"></param>
        private Text FindTextInColumn(BookmarkStart bookmark)
        {
            var cell = bookmark.GetParent<TableRow>().GetFirstChild<TableCell>();
            for (var i = 0; i < bookmark.ColumnFirst; i++)
                cell = cell.NextSibling<TableCell>();

            return cell.GetFirstDescendant<Text>();
        }
        /// <summary>
        /// Removes other OpenXML text elements in the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="keep"></param>
        private void RemoveOtherTexts(BookmarkStart bookmark, Text keep)
        {
            if (bookmark.ColumnFirst != null)
                return;

            Text text = null;
            var nextSibling = bookmark.NextSibling();
            while (text == null && nextSibling != null)
            {
                if (nextSibling.IsEndBookmark(bookmark))
                    break;
                foreach (var item in nextSibling.Descendants<Text>())
                {
                    if (item != keep)
                        item.Remove();
                }
                nextSibling = nextSibling.NextSibling();
            }
        }
        /// <summary>
        /// Inserts the text into the bookmark.
        /// </summary>
        /// <param name="bookmark"></param>
        /// <param name="value"></param>
        /// <param name="provider"></param>
        private void InsertBookmarkText(BookmarkStart bookmark, string value, IFormatProvider provider = default)
        {
            if (bookmark.NextSibling() is Run nextSubling)
            {
                var formFieldData = bookmark.Parent.Descendants<FormFieldData>().First(x => x.ChildElements.Any(el => el is FormFieldName formFieldName && formFieldName.Val.Value == bookmark.Name));
                var formFieldType = formFieldData?.GetFirstDescendant<TextBoxFormFieldType>();
                var format = formFieldData?.GetFirstDescendant<Format>();

                var text = value;
                if (formFieldType != null && format != null && formFieldType.Val.Value == TextBoxFormFieldValues.Date && !string.IsNullOrWhiteSpace(text))
                {
                    var datetime = DateTime.Parse(value, CultureInfo.InvariantCulture);
                    text = format.Val.HasValue
                        ? datetime.ToString(format.Val.Value, provider)
                        : value;
                }

                var run = new Run();
                var runProperties = nextSubling.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true);
                if (runProperties != null)
                    run.AppendChild(runProperties);
                run.AppendChild(new Text(text));

                var bookmarkEnd = bookmark.Parent.Descendants<BookmarkEnd>().First(x => x.Id == bookmark.Id);
                bookmark.Parent.InsertBefore(run, bookmarkEnd);
                RemoveOtherTexts(bookmark, run.GetFirstChild<Text>());
            }
        }
    }
}
