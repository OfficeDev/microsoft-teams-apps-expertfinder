// <copyright file="ResourceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Controllers
{
    using System;
    using System.Globalization;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Controller to handle strings.
    /// </summary>
    [Route("api/resource")]
    [ApiController]
    [Authorize]
    public class ResourceController : ControllerBase
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourceController"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public ResourceController(ILogger<ResourceController> logger, IStringLocalizer<Strings> localizer)
        {
            this.logger = logger;
            this.localizer = localizer;
        }

        /// <summary>
        /// Get resource strings for displaying in client app.
        /// </summary>
        /// <returns>Object containing required strings to be used in client app.</returns>
        [HttpGet]
        [Route("/api/resources/strings")]
        public ActionResult GetResourceStrings()
        {
            try
            {
                var strings = new
                {
                    Strings.SearchTextBoxPlaceholder,
                    Strings.InitialSearchResultMessageBodyText,
                    Strings.InitialSearchResultMessageHeaderText,
                    Strings.SearchResultNoItemsText,
                    Strings.SkillsTitle,
                    Strings.InterestTitle,
                    Strings.SchoolsTitle,
                    Strings.ViewButtonText,
                    Strings.MaxUserProfilesError,
                    Strings.UnauthorizedErrorMessage,
                    Strings.ForbiddenErrorMessage,
                    Strings.GeneralErrorMessage,
                    Strings.RefreshLinkText,
                };
                return this.Ok(strings);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting strings from resource controller.");
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}