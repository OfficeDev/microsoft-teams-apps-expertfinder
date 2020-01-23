// <copyright file="SearchRowResult.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    using System.Collections.Generic;

    /// <summary>
    /// Holds table row data from SharePoint search response data model.
    /// </summary>
    public class SearchRowResult
    {
        /// <summary>
        /// Gets or sets cell data from SharePoint search response table data model.
        /// </summary>
        public List<SearchPropertiesResult> Cells { get; set; }
    }
}
