// <copyright file="UserSearch.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    using System.Collections.Generic;

    /// <summary>
    /// User search request model.
    /// </summary>
    public class UserSearch
    {
        /// <summary>
        /// Gets or sets search text.
        /// </summary>
        public string SearchText { get; set; }

        /// <summary>
        /// Gets or sets search filters selected by user.
        /// </summary>
        public List<string> SearchFilters { get; set; }
    }
}