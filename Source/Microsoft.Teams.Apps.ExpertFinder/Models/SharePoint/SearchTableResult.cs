// <copyright file="SearchTableResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    using System.Collections.Generic;

    /// <summary>
    /// Holds table data from SharePoint search response data model.
    /// </summary>
    public class SearchTableResult
    {
        /// <summary>
        /// Gets or sets row collection from SharePoint search response data model.
        /// </summary>
        public List<SearchRowResult> Rows { get; set; }
    }
}
