// <copyright file="TokenSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.Configuration
{
    /// <summary>
    /// Provides app setting related to jwt token.
    /// </summary>
    public class TokenSettings : AADSettings
    {
        /// <summary>
        /// Gets application base uri.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets random key to create jwt security key.
        /// </summary>
        public string SecurityKey { get; set; }
    }
}
