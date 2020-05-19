// <copyright file="GraphApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// The class that represent the helper methods to access Microsoft Graph API.
    /// </summary>
    public class GraphApiHelper : IGraphApiHelper
    {
        /// <summary>
        /// Post user details to API request url.
        /// </summary>
        private const string UserProfileGraphEndpointUrl = "https://graph.microsoft.com/v1.0/me";

        /// <summary>
        /// Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.
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

        /// <inheritdoc/>
        public async Task<bool> UpdateUserProfileDetailsAsync(string token, string body)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Patch, UserProfileGraphEndpointUrl))
            {
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                using (var userProfileUpdateResponse = await this.client.SendAsync(request).ConfigureAwait(false))
                {
                    if (userProfileUpdateResponse.IsSuccessStatusCode)
                    {
                        return true;
                    }

                    var errorMessage = await userProfileUpdateResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                    this.logger.LogInformation($"Graph API user profile update error: {errorMessage}");

                    return false;
                }
            }
        }

        /// <inheritdoc/>
        public async Task<UserProfileDetail> GetUserProfileAsync(string token)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Get, $"{UserProfileGraphEndpointUrl}?$select=id,displayname,jobTitle,aboutme,skills,interests,schools"))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                using (var response = await this.client.SendAsync(request).ConfigureAwait(false))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        return JsonConvert.DeserializeObject<UserProfileDetail>(json);
                    }

                    var errorMessage = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    this.logger.LogInformation($"Error getting user profile from Graph: {errorMessage}");

                    return null;
                }
            }
        }
    }
}