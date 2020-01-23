// <copyright file="SharePointApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Extensions;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Handles api calls for SharePoint to get user details based on query.
    /// </summary>
    public class SharePointApiHelper : ISharePointApiHelper
    {
        /// <summary>
        /// Default SharePoint search filter criteria.
        /// </summary>
        private const string DefaultSearchType = "skills";

        /// <summary>
        /// SharePoint constant source id for user profile search.
        /// </summary>
        private const string SharePointSearchSourceId = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";

        /// <summary>
        /// Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.
        /// </summary>
        /// </summary>
        private readonly HttpClient client;

        /// <summary>
        /// Initializes a new instance of the <see cref="SharePointApiHelper"/> class.
        /// Helper class to handle SharePoint search rest api to get user profiles based on search criteria.
        /// </summary>
        /// <param name="client">Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.</param>
        public SharePointApiHelper(HttpClient client)
        {
            this.client = client;
        }

        /// <summary>
        /// Get user profiles from SharePoint based on search text and filters.
        /// </summary>
        /// <param name="searchText">Search text to match.</param>
        /// <param name="searchFilters">List of property filters to perform serch on.</param>
        /// <param name="token">SharePoint user access token.</param>
        /// <param name="resourceBaseUrl">SharePoint base uri.</param>
        /// <returns>User profile collection that matches search query.</returns>
        public async Task<IList<UserProfileDetail>> GetUserProfilesAsync(string searchText, IList<string> searchFilters, string token, string resourceBaseUrl)
        {
            List<UserProfileDetail> userDetails = new List<UserProfileDetail>();

            var requestUrl = this.GetSharePointSearchRequestUri(searchText, searchFilters, resourceBaseUrl);
            HttpMethod httpMethod = new HttpMethod("GET");
            var request = new HttpRequestMessage(httpMethod, requestUrl);

            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await this.client.SendAsync(request).ConfigureAwait(false);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                var searchResult = JsonConvert.DeserializeObject<SearchResponse>(JObject.Parse(result).SelectToken("PrimaryQueryResult").ToString());
                var searchResultRows = searchResult.RelevantResults.Table.Rows;

                userDetails = searchResultRows.Select(user => new UserProfileDetail()
                {
                    AboutMe = user.Cells.GetCellsValue("AboutMe"),
                    Interests = user.Cells.GetCellsValue("Interests"),
                    JobTitle = user.Cells.GetCellsValue("JobTitle"),
                    PreferredName = user.Cells.GetCellsValue("PreferredName"),
                    Schools = user.Cells.GetCellsValue("Schools"),
                    Skills = user.Cells.GetCellsValue("Skills"),
                    WorkEmail = user.Cells.GetCellsValue("WorkEmail"),
                    Path = user.Cells.GetCellsValue("OriginalPath"),
                }).ToList();
            }
            else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                var errorMessage = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                throw new UnauthorizedAccessException($"Invalid user access token, error - {errorMessage}");
            }
            else
            {
                var errorMessage = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                throw new Exception($"{response.ReasonPhrase} - {errorMessage}");
            }

            return userDetails;
        }

        /// <summary>
        /// Generate SharePoint search query rest api uri.
        /// </summary>
        /// <param name="searchText">Search text to match.</param>
        /// <param name="searchFilters">List of property filters to perform serch on.</param>
        /// <param name="baseUri">SharePoint base uri.</param>
        /// <returns>SharePoint search query rest api uri.</returns>
        /// <remark>Returned url will be like "https://{SharepointSteName}.sharepoint.com/_api/search/query?querytext='{SearchQuery}'&sourceid=B09A7990-05EA-4AF9-81EF-EDFAB16C4E31".
        private string GetSharePointSearchRequestUri(string searchText, IList<string> searchFilters, string baseUri)
        {
            StringBuilder searchString = new StringBuilder();

            if (searchFilters != null && searchFilters.Count > 0)
            {
                if (searchFilters.Count > 1)
                {
                    var items = searchFilters.Take(searchFilters.Count - 1).ToList();
                    items.ForEach(value =>
                    {
                        searchString.Append(value + ":" + searchText + " OR ");
                    });
                }

                searchString.Append(searchFilters.Last() + ":" + searchText);
            }
            else
            {
                searchString.Append(DefaultSearchType + ":" + searchText);
            }

            return $"{baseUri}_api/search/query?querytext='{searchString.ToString()}'&sourceid='{SharePointSearchSourceId}'";
        }
    }
}
