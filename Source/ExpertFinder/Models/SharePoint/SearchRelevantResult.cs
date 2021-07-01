// <copyright file="SearchRelevantResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    /// <summary>
    /// Holds relevant result data from from SharePoint search response data model.
    /// </summary>
    public class SearchRelevantResult
    {
        /// <summary>
        /// Gets or sets table data from SharePoint search response data model.
        /// </summary>
        public SearchTableResult Table { get; set; }
    }
}
