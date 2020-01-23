// <copyright file="UserProfileDetailBase.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Model for submit action on edit profile for microsoft graph api.
    /// </summary>
    public class UserProfileDetailBase
    {
        /// <summary>
        /// Gets or sets user about me.
        /// </summary>
        [JsonProperty("aboutMe")]
        public string AboutMe { get; set; }

        /// <summary>
        /// Gets or sets skill details.
        /// </summary>
        [JsonProperty("skills")]
        public List<string> Skills { get; set; }

        /// <summary>
        /// Gets or sets interest details.
        /// </summary>
        [JsonProperty("interests")]
        public List<string> Interests { get; set; }

        /// <summary>
        /// Gets or sets school details.
        /// </summary>
        [JsonProperty("schools")]
        public List<string> Schools { get; set; }
    }
}