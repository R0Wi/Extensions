﻿using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WebAPI.Helpers
{
    /// <summary>
    /// The Open XML extensions methods.
    /// </summary>
    public static class OpenXMLElementExtensions
    {
        /// <summary>
        /// Gets the first descendant.
        /// </summary>
        /// <typeparam name="T">The element type.</typeparam>
        /// <param name="parent"></param>
        /// <returns>If a descendant is found return it, otherwise, return null.</returns>
        /// <exception cref="ArgumentNullException">The <see cref="OpenXmlElement"/> is null.</exception>
        public static T GetFirstDescendant<T>(this OpenXmlElement parent) where T : OpenXmlElement
        {
            if (parent == null)
                throw new ArgumentNullException(nameof(parent));

            return parent.Descendants<T>()?.FirstOrDefault();
        }
        /// <summary>
        /// Gets the parent.
        /// </summary>
        /// <typeparam name="T">The element type.</typeparam>
        /// <param name="child"></param>
        /// <returns>if a parent is found return it, otherwise, return null.</returns>
        /// <exception cref="ArgumentNullException">The <see cref="OpenXmlElement"/> is null.</exception>
        public static T GetParent<T>(this OpenXmlElement child) where T : OpenXmlElement
        {
            if (child == null)
                throw new ArgumentNullException(nameof(child));

            while (child != null)
            {
                child = child.Parent;
                if (child is T)
                    return (T)child;
            }
            return null;
        }
        /// <summary>
        /// Is the end of the bookmark.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="startBookmark">The start of the bookmark element.</param>
        /// <returns>If it is the end of the bookmark return true. If element is not <see cref="BookmarkEnd"/> or it is not the end of the bookmark return false.</returns>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        public static bool IsEndBookmark(this OpenXmlElement element, BookmarkStart startBookmark)
        {
            if (startBookmark == null)
                throw new ArgumentNullException(nameof(startBookmark));

            return IsEndBookmark(element as BookmarkEnd, startBookmark);
        }
        /// <summary>
        /// Is the end of the bookmark.
        /// </summary>
        /// <param name="endBookmark">The end of the bookmark element.</param>
        /// <param name="startBookmark">The start of the bookmark element.</param>
        /// <exception cref="ArgumentNullException">The <see cref="BookmarkStart"/> is null.</exception>
        /// <returns>If it is the end of the bookmark return true. If element is not <see cref="BookmarkEnd"/> or it is not the end of the bookmark return false.</returns>
        public static bool IsEndBookmark(this BookmarkEnd endBookmark, BookmarkStart startBookmark)
        {
            if (startBookmark == null)
                throw new ArgumentNullException(nameof(startBookmark));

            return endBookmark == null
                ? false
                : endBookmark.Id.HasValue && startBookmark.Id.HasValue && (endBookmark.Id.Value == startBookmark.Id.Value);
        }
    }
}
