using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml
{
    /// <summary>
    /// Represents the word document editor.
    /// </summary>
    public static class WordDocumentEditor
    {
        /// <summary>
        /// Builds the document from the template with bookmarks.
        /// </summary>
        /// <param name="content">The docx template file.</param>
        /// <param name="bookmarkValues">The bookmark values.</param>
        /// <param name="provider">The culture info.</param>
        /// <exception cref="ArgumentNullException">Thrown when one of the parameters: content or bookmarkValues is null or empty.</exception>
        /// <exception cref="OpenXmlPackageException">Thrown when the package is not valid Open XML WordprocessingDocument.</exception>
        /// <returns>The content with mimetype: application/vnd.openxmlformats-officedocument.wordprocessingml.document</returns>
        public static byte[] SetDocumentBookmarkValues(byte[] content, IDictionary<string, string> bookmarkValues, IFormatProvider provider = default)
        {
            if (content == null || !content.Any())
                throw new ArgumentNullException(nameof(content));
            if (bookmarkValues == null || !bookmarkValues.Any())
                throw new ArgumentNullException(nameof(bookmarkValues));

            byte[] fileContents = null;
            using (var stream = new MemoryStream(content, true))
            {
                using (var wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    OpenXmlBookmarkHelper.SetDocumentBookmarkValues(wordDocument, bookmarkValues, provider);
                }
                fileContents = stream.ToArray();
            }
            return fileContents;
        }
        /// <summary>
        /// Gets the all document bookmarks values.
        /// </summary>
        /// <param name="content">The docx template file.</param>
        /// <param name="provider">The culture info.</param>
        /// <exception cref="ArgumentNullException">Thrown when one of the parameters: content or bookmarkValues is null or empty.</exception>
        /// <exception cref="OpenXmlPackageException">Thrown when the package is not valid Open XML WordprocessingDocument.</exception>
        public static IDictionary<string, string> GetDocumentBookmarkValues(byte[] content, IFormatProvider provider = default)
        {
            if (content == null || !content.Any())
                throw new ArgumentNullException(nameof(content));

            IDictionary<string, string> result;
            using (var stream = new MemoryStream(content, false))
            {
                using var wordDocument = WordprocessingDocument.Open(stream, false);
                result = OpenXmlBookmarkHelper.GetDocumentBookmarkValues(wordDocument, false, provider);
            }
            return result;
        }
    }
}
