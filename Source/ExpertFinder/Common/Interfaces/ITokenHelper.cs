// <copyright file="ITokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces
{
    using System.Threading.Tasks;

    /// <summary>
    /// Helper class to generate Azure Active Directory user access token for given resource, e.g. Microsoft Graph.
    /// </summary>
    public interface ITokenHelper
    {
        /// <summary>
        /// Get user access token for given resource using Bot OAuth client instance.
        /// </summary>
        /// <param name="fromId">Activity from id.</param>
        /// <param name="resourceUrl">Resource url for which token will be acquired.</param>
        /// <returns>A task that represents security access token for given resource.</returns>
        Task<string> GetUserTokenAsync(string fromId, string resourceUrl);
    }
}