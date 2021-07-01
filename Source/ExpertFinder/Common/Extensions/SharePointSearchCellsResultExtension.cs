// <copyright file="SharePointSearchCellsResultExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common.Extensions
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;

    /// <summary>
    /// A class that extends <see cref="SearchPropertiesResult"/> class to get given key value.
    /// </summary>
    public static class SharePointSearchCellsResultExtension
    {
        /// <summary>
        /// Get value assosciated with key from SharePoint search response cell data model.
        /// </summary>
        /// <param name="cells"> Collection of data from SharePoint search response cell data model.</param>
        /// <param name="key">String key that is used to get value associated with.</param>
        /// <returns>Value that matches given key in provided collection of SharePoint search response cell data.</returns>
        public static string GetCellsValue(this IEnumerable<SearchPropertiesResult> cells, string key)
        {
            return cells.Where(item => item.Key == key).Select(item => item.Value).FirstOrDefault();
        }
    }
}
