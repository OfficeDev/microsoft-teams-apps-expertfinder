// <copyright file="UserProfileDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// User profile details model class.
    /// </summary>
    public class UserProfileDetail : UserProfileDetailBase
    {
        /// <summary>
        /// Gets or sets odataContext.
        /// </summary>
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }

        /// <summary>
        /// Gets or sets user unique id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets user display name.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets user job title.
        /// </summary>
        [JsonProperty("jobTitle")]
        public string JobTitle { get; set; }
    }
}