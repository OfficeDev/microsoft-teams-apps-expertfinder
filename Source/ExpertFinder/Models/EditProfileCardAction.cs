// <copyright file="EditProfileCardAction.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Edit profile task module model class.
    /// </summary>
    public class EditProfileCardAction
    {
        /// <summary>
        /// Gets or sets msteams card action type.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction Msteams { get; set; }

        /// <summary>
        /// Gets or sets bot command name.
        /// </summary>
        [JsonProperty("command")]
        public string Command { get; set; }

        /// <summary>
        /// Gets or sets user profile card unique id.
        /// </summary>
        [JsonProperty("MyProfileCardId")]
        public string MyProfileCardId { get; set; }

        /// <summary>
        /// Gets or sets user about me details.
        /// </summary>
        [JsonProperty("aboutMe")]
        public string AboutMe { get; set; }

        /// <summary>
        /// Gets or sets user interest details.
        /// </summary>
        [JsonProperty("interests")]
        public string Interests { get; set; }

        /// <summary>
        /// Gets or sets user school details.
        /// </summary>
        [JsonProperty("schools")]
        public string Schools { get; set; }

        /// <summary>
        /// Gets or sets user skill details.
        /// </summary>
        [JsonProperty("skills")]
        public string Skills { get; set; }
    }
}