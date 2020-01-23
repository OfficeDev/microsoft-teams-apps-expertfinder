// <copyright file="BotSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.Configuration
{
    /// <summary>
    /// Provides app settings related to Expert Finder bot.
    /// </summary>
    public class BotSettings : AADSettings
    {
        /// <summary>
        /// Gets or sets application base uri.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets SharePoint site Uri.
        /// </summary>
        public string SharePointSiteUrl { get; set; }

        /// <summary>
        /// Gets or sets application Insights instrumentation key which we passes to client application.
        /// </summary>
        public string AppInsightsInstrumentationKey { get; set; }

        /// <summary>
        /// Gets or sets tenant id.
        /// </summary>
        public string TenantId { get; set; }
    }
}
