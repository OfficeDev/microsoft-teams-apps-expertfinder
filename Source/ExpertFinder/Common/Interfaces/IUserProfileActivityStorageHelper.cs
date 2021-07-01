// <copyright file="IUserProfileActivityStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ExpertFinder.Models;

    /// <summary>
    /// Implements storage helper which stores user profile card activity details in Microsoft Azure Table service.
    /// </summary>
    public interface IUserProfileActivityStorageHelper
    {
        /// <summary>
        /// Stores or update user profile card activity id and user profile card id in table storage.
        /// </summary>
        /// <param name="userProfileConversatioEntity">Holds user profile activity id and card id to uniquely identify user activity that is being edited.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents user profile activity information is saved or updated.False indicates failure in saving data. </returns>
        Task<bool> UpsertUserProfileConversationDataAsync(UserProfileActivityInfo userProfileConversatioEntity);

        /// <summary>
        /// Get user profile card activity id and user profile card id from table storage based on user profile card id.
        /// </summary>
        /// <param name="myProfileCardId">Unique user profile card id.</param>
        /// <returns>A task that represent object to hold user profile card activity id and user profile card id.</returns>
        Task<UserProfileActivityInfo> GetUserProfileConversationDataAsync(string myProfileCardId);
    }
}