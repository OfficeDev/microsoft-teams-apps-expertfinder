// <copyright file="SearchResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    /// <summary>
    /// SharePoint Search api response data model.
    /// </summary>
    public class SearchResponse
    {
        /// <summary>
        /// Gets or sets relevant results from SharePoint search response data.
        /// </summary>
        public SearchRelevantResult RelevantResults { get; set; }
    }
}
