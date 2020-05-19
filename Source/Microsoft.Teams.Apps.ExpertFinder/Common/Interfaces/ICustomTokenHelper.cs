// <copyright file="ICustomTokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System.Collections.Generic;
    using System.Security.Claims;

    /// <summary>
    /// Helper class for JWT token generation and validation for given resource, e.g. SharePoint.
    /// </summary>
    public interface ICustomTokenHelper
    {
        /// <summary>
        /// Generate custom jwt access token to authenticate/verify valid request on API side.
        /// </summary>
        /// <param name="aadObjectId">User account's object id within Azure Active Directory.</param>
        /// <param name="serviceURL">Service uri where responses to this activity should be sent.</param>
        /// <param name="fromId">Unique user id from activity.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token.</param>
        /// <returns>Custom jwt access token.</returns>
        string GenerateAPIAuthToken(string aadObjectId, string serviceURL, string fromId, int jwtExpiryMinutes);
    }
}