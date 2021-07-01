// <copyright file="IGraphApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ExpertFinder.Models;

    /// <summary>
    /// Provides the helper methods to access Microsoft Graph API.
    /// </summary>
    public interface IGraphApiHelper
    {
        /// <summary>
        /// Get user profile details from Microsoft Graph.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <returns>User profile details.</returns>
        Task<UserProfileDetail> GetUserProfileAsync(string token);

        /// <summary>
        /// Call Microsoft Graph API to update user profile details.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="body">User profile.</param>
        /// <returns>A task that returns true if user profile is successfully updated and false if it fails.</returns>
        /// <remark>Reference link for Graph API used for updating user profile is"https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-beta&amp;tabs=http".</remark>
        Task<bool> UpdateUserProfileDetailsAsync(string token, string body);
    }
}