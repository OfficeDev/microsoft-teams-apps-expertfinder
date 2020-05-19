// <copyright file="ISharePointApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;

    /// <summary>
    /// Handles API calls for SharePoint to get user details based on query.
    /// </summary>
    public interface ISharePointApiHelper
    {
        /// <summary>
        /// Get user profiles from SharePoint based on search text and filters.
        /// </summary>
        /// <param name="searchText">Search text to match.</param>
        /// <param name="searchFilters">List of property filters to perform serch on.</param>
        /// <param name="token">SharePoint user access token.</param>
        /// <param name="resourceBaseUrl">SharePoint base uri.</param>
        /// <returns>User profile collection that matches search query.</returns>
        Task<IList<UserProfileDetail>> GetUserProfilesAsync(string searchText, IList<string> searchFilters, string token, string resourceBaseUrl);
    }
}
