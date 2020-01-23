// <copyright file="UserProfileActivityInfo.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Holds user profile activity id and card id to uniquely identify user activity that is being edited.
    /// </summary>
    public class UserProfileActivityInfo : TableEntity
    {
        /// <summary>
        /// Partition key for UserProfile table.
        /// </summary>
        public const string UserProfileActivityInfoPartitionKey = "UserProfileActivityInfo";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserProfileActivityInfo"/> class.
        /// Holds user profile activity id and card id to uniquely identify user activity that is being edited.
        /// </summary>
        public UserProfileActivityInfo()
        {
            this.PartitionKey = UserProfileActivityInfoPartitionKey;
        }

        /// <summary>
        /// Gets or sets user profile card activity id.
        /// </summary>
        public string MyProfileCardActivityId { get; set; }

        /// <summary>
        /// Gets or sets custom unique guid id of user profile card.
        /// </summary>
        public string MyProfileCardId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }
    }
}