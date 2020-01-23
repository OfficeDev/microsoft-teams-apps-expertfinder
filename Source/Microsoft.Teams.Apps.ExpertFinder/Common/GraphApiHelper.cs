// <copyright file="GraphApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// The class that represent the helper methods to access Microsoft Graph api.
    /// </summary>
    public class GraphApiHelper : IGraphApiHelper
    {
        /// <summary>
        /// Post user details to api request url.
        /// </summary>
        private const string UserProfileGraphEndpointUrl = "https://graph.microsoft.com/v1.0/me";

        /// <summary>
        /// Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.
        /// </summary>
        /// </summary>
        private readonly HttpClient client;

        /// <summary>
        /// Instance to send logs to the Application Insights service..
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphApiHelper"/> class.
        /// </summary>
        /// <param name="client">Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public GraphApiHelper(HttpClient client, ILogger<GraphApiHelper> logger)
        {
            this.client = client;
            this.logger = logger;
        }

        /// <summary>
        /// Call Microsoft Graph api to work with user profile details.
        /// </summary>
        /// <param name="token">Microsoft Graph api user access token.</param>
        /// <param name="body">User graph api request body.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        /// <remark>Reference link for Graph api used for updating user profile is"https://docs.microsoft.com/en-us/graph/api/user-update?view=graph-rest-beta&tabs=http".</remark>
        public async Task<bool> UpdateUserProfileDetailsAsync(string token, string body)
        {
            string requestUrl = UserProfileGraphEndpointUrl;
            HttpMethod httpMethod = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(httpMethod, requestUrl);

            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            request.Content = new StringContent(body, Encoding.UTF8, "application/json");
            var userProfileUpdateResponse = await this.client.SendAsync(request).ConfigureAwait(false);

            if (userProfileUpdateResponse.IsSuccessStatusCode)
            {
                return true;
            }

            var errorMessage = await userProfileUpdateResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
            this.logger.LogInformation($"Graph api user profile update error- {errorMessage}");
            return false;
        }

        /// <summary>
        /// Get user profile details from Microsoft Graph.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <returns>User profile details.</returns>
        public async Task<UserProfileDetail> GetUserProfileAsync(string token)
        {
            var response = await this.GetUserProfileDetailsAsync(token, $"{UserProfileGraphEndpointUrl}?$select=id,displayname,jobTitle,aboutme,skills,interests,schools").ConfigureAwait(false);

            if (response.IsSuccessStatusCode)
            {
                return await this.DeserializeJsonStringAsync<UserProfileDetail>(response).ConfigureAwait(false);
            }

            var errorMessage = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            this.logger.LogInformation($"Graph api getting user profile error- {errorMessage}");
            return null;
        }

        /// <summary>
        /// Get user profile detail from Microsoft Graph api.
        /// </summary>
        /// <param name="token">Microsoft Graph api user access token.</param>
        /// <param name="requestUrl">Microsoft Graph user profile request api uri.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        private async Task<HttpResponseMessage> GetUserProfileDetailsAsync(string token, string requestUrl)
        {
            HttpMethod httpMethod = new HttpMethod("GET");
            var request = new HttpRequestMessage(httpMethod, requestUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            return await this.client.SendAsync(request).ConfigureAwait(false);
        }

        /// <summary>
        /// Deserialize the response from HTTP response data.
        /// </summary>
        /// <typeparam name="T">Model to deserialize data.</typeparam>
        /// <param name="response">Represents a HTTP response message including the status code and data.</param>
        /// <returns>Deserialized HTTP response data.</returns>
        private async Task<T> DeserializeJsonStringAsync<T>(HttpResponseMessage response)
            where T : class
        {
            return JsonConvert.DeserializeObject<T>(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
        }
    }
}