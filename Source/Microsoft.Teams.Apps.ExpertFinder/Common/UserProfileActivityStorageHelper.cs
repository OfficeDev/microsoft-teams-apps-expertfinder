// <copyright file="UserProfileActivityStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage helper which stores user profile card activity details in Microsoft Azure Table service.
    /// </summary>
    public class UserProfileActivityStorageHelper : IUserProfileActivityStorageHelper
    {
        /// <summary>
        /// Task for initialization.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Microsoft Azure Table Storage connection string.
        /// </summary>
        private readonly string connectionString;

        /// <summary>
        /// Microsoft Azure Table Storage table name.
        /// </summary>
        private readonly string tableName;

        /// <summary>
        /// Represents a table in the Microsoft Azure Table service.
        /// </summary>
        private CloudTable profileCloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserProfileActivityStorageHelper"/> class.
        /// Handles Microsoft Azure Table read write operations.
        /// </summary>
        /// <param name="configuration">Object that passes the application configuration key-values.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for Azure Table storage.</param>
        public UserProfileActivityStorageHelper(IOptionsMonitor<StorageSettings> optionsAccessor)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync());
            this.connectionString = optionsAccessor.CurrentValue.StorageConnectionString;
            this.tableName = "UserProfileActivityInfo";
        }

        /// <summary>
        /// Stores or update user profile card activity id and user profile card id in table storage.
        /// </summary>
        /// <param name="userProfileConversationEntity">Holds user profile activity id and card id to uniquely identify user activity that is being edited.</param>
        /// <returns>A task that represents user profile conversation data is saved or updated.</returns>
        public async Task<bool> UpsertConverationStateAsync(UserProfileActivityInfo userProfileConversationEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(userProfileConversationEntity).ConfigureAwait(false);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get user profile card activity id and user profile card id from table storage based on user profile card id.
        /// </summary>
        /// <param name="myProfileCardId">Unique user profile card id.</param>
        /// <returns>A task that represent object to hold user profile card activity id and user profile card id.</returns>
        public async Task<UserProfileActivityInfo> GetUserProfileConversationDataAsync(string myProfileCardId)
        {
            TableResult searchResult;
            await this.EnsureInitializedAsync().ConfigureAwait(false);
            var searchOperation = TableOperation.Retrieve<UserProfileActivityInfo>(UserProfileActivityInfo.UserProfileActivityInfoPartitionKey, myProfileCardId);
            searchResult = await this.profileCloudTable.ExecuteAsync(searchOperation).ConfigureAwait(false);
            return (UserProfileActivityInfo)searchResult.Result;
        }

        /// <summary>
        /// Store or update user profile activity information entity which holds user profile card activity id and user profile card id in table storage.
        /// </summary>
        /// <param name="entity">Object that contains user profile card activity id and user profile card unique id.</param>
        /// <returns>A task that represents configuration entity is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(UserProfileActivityInfo entity)
        {
            await this.EnsureInitializedAsync().ConfigureAwait(false);
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
            return await this.profileCloudTable.ExecuteAsync(addOrUpdateOperation).ConfigureAwait(false);
        }

        /// <summary>
        /// Create UserProfile table if it doesnt exists.
        /// </summary>
        /// <returns>A<see cref="Task"/> representing the asynchronous operation task which represents table is created if its not exists.</returns>
        private async Task InitializeAsync()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(this.connectionString);
            CloudTableClient cloudTableClient = storageAccount.CreateCloudTableClient();
            this.profileCloudTable = cloudTableClient.GetTableReference(this.tableName);
            await this.profileCloudTable.CreateIfNotExistsAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Ensures .Microsoft Azure Table Storage should be created before working on table.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value.ConfigureAwait(false);
        }
    }
}