// <copyright file="UserProfileController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Controllers
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;

    /// <summary>
    /// Controller to handle SharePoint API operations.
    /// </summary>
    [Route("api/users")]
    [ApiController]
    [Authorize]
    public class UserProfileController : ControllerBase
    {
        /// <summary>
        /// Helper for acquiring AAD token for given resource.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Instance of SharePoint search REST API helper.
        /// </summary>
        private readonly ISharePointApiHelper sharePointApiHelper;

        /// <summary>
        /// SharePoint site uri.
        /// </summary>
        private readonly string sharePointSiteUri;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserProfileController"/> class.
        /// </summary>
        /// <param name="sharePointApiHelper">Instance of SharePoint search REST API helper.</param>
        /// <param name="tokenHelper">Instance of class for validating custom jwt access token.</param>
        /// <param name="botSettings">A set of key/value application configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public UserProfileController(ISharePointApiHelper sharePointApiHelper, ITokenHelper tokenHelper, IOptionsMonitor<BotSettings> botSettings, ILogger<UserProfileController> logger)
        {
            this.sharePointApiHelper = sharePointApiHelper;
            this.tokenHelper = tokenHelper;
            this.sharePointSiteUri = botSettings.CurrentValue.SharePointSiteUrl;
            this.logger = logger;
        }

        /// <summary>
        /// Post call to search service.
        /// </summary>
        /// <param name="searchQuery">User search query which includes search text and search filters.</param>
        /// <returns>List of user profile details which matches search text for properties given by search filters.</returns>
        public async Task<IActionResult> Post(UserSearch searchQuery)
        {
            try
            {
                var jwtToken = this.Request.Headers["Authorization"].ToString().Split(' ')[1];

                if (searchQuery == null)
                {
                    return this.StatusCode(StatusCodes.Status403Forbidden);
                }

                var fromId = this.User.Claims.Where(claim => claim.Type == "fromId").Select(claim => claim.Value).FirstOrDefault();
                if (string.IsNullOrEmpty(fromId))
                {
                    this.logger.LogInformation("Failed to get fromId from token.");
                    return this.StatusCode(StatusCodes.Status401Unauthorized);
                }

                var userToken = await this.tokenHelper.GetUserTokenAsync(fromId, this.sharePointSiteUri).ConfigureAwait(false);
                this.logger.LogInformation("Initiated call to user search service");
                var userProfiles = await this.sharePointApiHelper.GetUserProfilesAsync(searchQuery.SearchText, searchQuery.SearchFilters, userToken, this.sharePointSiteUri).ConfigureAwait(false);

                this.logger.LogInformation("Call to search service succeeded");
                return this.Ok(userProfiles);
            }
            catch (UnauthorizedAccessException ex)
            {
                this.logger.LogError(ex, "Failed to get user token to make post call to api.");
                return this.StatusCode(StatusCodes.Status401Unauthorized);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making post call to search service.");
                return this.BadRequest(ex.Message);
            }
        }
    }
}