// <copyright file="UserProfileDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint
{
    /// <summary>
    /// User details model for SharePoint to show as api request response.
    /// </summary>
    public class UserProfileDetail
    {
        /// <summary>
        /// Gets or sets user about me.
        /// </summary>
        public string AboutMe { get; set; }

        /// <summary>
        /// Gets or sets user interest.
        /// </summary>
        public string Interests { get; set; }

        /// <summary>
        /// Gets or sets user job title.
        /// </summary>
        public string JobTitle { get; set; }

        /// <summary>
        /// Gets or sets user schools.
        /// </summary>
        public string Schools { get; set; }

        /// <summary>
        /// Gets or sets user name.
        /// </summary>
        public string PreferredName { get; set; }

        /// <summary>
        /// Gets or sets user skills.
        /// </summary>
        public string Skills { get; set; }

        /// <summary>
        /// Gets or sets user work email.
        /// </summary>
        public string WorkEmail { get; set; }

        /// <summary>
        /// Gets or sets user picture path.
        /// </summary>
        public string Path { get; set; }
    }
}
