// <copyright file="BotSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.Configuration
{
    /// <summary>
    /// Provides app settings related to Expert Finder bot.
    /// </summary>
    public class BotSettings
    {
        /// <summary>
        /// Gets or sets application base uri.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets application Insights instrumentation key which we passes to client application.
        /// </summary>
        public string AppInsightsInstrumentationKey { get; set; }

        /// <summary>
        /// Gets or sets bot OAuth connection name.
        /// </summary>
        public string OAuthConnectionName { get; set; }

        /// <summary>
        /// Gets or sets a random key used to sign the JWT sent to the task module.
        /// </summary>
        public string TokenSigningKey { get; set; }

        /// <summary>
        /// Gets or sets SharePoint site Uri.
        /// </summary>
        public string SharePointSiteUrl { get; set; }

        /// <summary>
        /// Gets or sets Azure Table Storage connection string.
        /// </summary>
        public string StorageConnectionString { get; set; }

        /// <summary>
        /// Gets or sets tenant id.
        /// </summary>
        public string TenantId { get; set; }
    }
}
