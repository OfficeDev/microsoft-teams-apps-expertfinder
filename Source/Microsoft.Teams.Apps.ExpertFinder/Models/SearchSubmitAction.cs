// <copyright file="SearchSubmitAction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;
    using Newtonsoft.Json;

    /// <summary>
    /// Submit action on view of search task module model class.
    /// </summary>
    public class SearchSubmitAction
    {
        /// <summary>
        /// Gets or sets commands from which task module is invoked.
        /// </summary>
        [JsonProperty("command")]
        public string Command { get; set; }

        /// <summary>
        /// Gets or sets user profile details from task module.
        /// </summary>
        [JsonProperty("searchresults")]
        public List<SharePoint.UserProfileDetail> UserProfiles { get; set; }
    }
}
