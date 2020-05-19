// <copyright file="MainDialog.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.ExpertFinder.Cards;
    using Microsoft.Teams.Apps.ExpertFinder.Common;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Dialogs;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Main Dialog class.
    /// </summary>
    public class MainDialog : LogoutDialog
    {
        /// <summary>
        /// Invoke activity type.
        /// </summary>
        private const string InvokeActivityType = "invoke";

        /// <summary>
        /// Message activity type.
        /// </summary>
        private const string MessageActivityType = "message";

        /// <summary>
        /// Sign in verify activity type.
        /// </summary>
        private const string SignInActivityName = "signin/verifyState";

        /// <summary>
        /// Helper for working with Microsoft Graph API.
        /// </summary>
        private readonly IGraphApiHelper graphApiHelper;

        /// <summary>
        /// Helper for working with Microsoft Azure Table storage service.
        /// </summary>
        private readonly IUserProfileActivityStorageHelper storageHelper;

        /// <summary>
        /// Instance to send logs to the Application Insights service..
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainDialog"/> class.
        /// </summary>
        /// <param name="graphApiHelper">Helper for working with Microsoft Graph API.</param>
        /// <param name="storageHelper">Helper for working with Microsoft Azure Table storage service.</param>
        /// <param name="botSettings">A set of key/value application configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public MainDialog(IGraphApiHelper graphApiHelper, IUserProfileActivityStorageHelper storageHelper, IOptionsMonitor<BotSettings> botSettings, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), botSettings.CurrentValue.OAuthConnectionName)
        {
            this.graphApiHelper = graphApiHelper;
            this.storageHelper = storageHelper;
            this.logger = logger;

            this.AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = botSettings.CurrentValue.OAuthConnectionName,
                    Text = Strings.SignInCardText,
                    Title = Strings.SignInButtonText,
                    Timeout = (int)TimeSpan.FromMinutes(5).TotalMilliseconds,
                }));

            this.AddDialog(new WaterfallDialog(
                "MainDialog",
                new WaterfallStep[] { this.CheckForUnknownInputAsync, this.OAuthPromptStepAsync, this.MyProfileAndSearchAsync }));

            // The initial child Dialog to run.
            this.InitialDialogId = "MainDialog";
        }

        /// <summary>
        /// Check for unknown input, and show the help prompt.
        /// </summary>
        /// <param name="stepContext">Provides context for a step in a bot dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Tracking task.</returns>
        private async Task<DialogTurnResult> CheckForUnknownInputAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var activity = stepContext.Context.Activity;
            if (activity.Type == ActivityTypes.Message)
            {
                switch (activity.Text?.Trim())
                {
                    case Constants.MyProfile:
                    case Constants.Search:
                    case null:
                        return await stepContext.NextAsync();

                    default:
                        await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(HelpCard.GetHelpCard())).ConfigureAwait(false);
                        return await stepContext.EndDialogAsync().ConfigureAwait(false);
                }
            }

            return await stepContext.NextAsync();
        }

        /// <summary>
        /// Initiate prompt for user sign-in.
        /// </summary>
        /// <param name="stepContext">Provides context for a step in a bot dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Tracking task.</returns>
        private async Task<DialogTurnResult> OAuthPromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var activity = stepContext.Context.Activity;

            stepContext.Values["command"] = activity.Text?.Trim();
            if (activity.Text == null && activity.Value != null && activity.Type == MessageActivityType)
            {
                stepContext.Values["command"] = JToken.Parse(activity.Value.ToString()).SelectToken("command").ToString();
            }

            this.logger.LogInformation($"Sent sign-in card for conversation id: {activity.Conversation.Id}");
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Get user profile or search or edit profile based on activity type.
        /// </summary>
        /// <param name="stepContext">Provides context for a step in a bot dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Tracking task.</returns>
        private async Task<DialogTurnResult> MyProfileAndSearchAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var activity = stepContext.Context.Activity;

            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse == null)
            {
                this.logger.LogInformation($"User is not authenticated and token is null for: {activity.Conversation.Id}.");
                await stepContext.Context.SendActivityAsync(Strings.NotLoggedInText).ConfigureAwait(false);
                return await stepContext.EndDialogAsync().ConfigureAwait(false);
            }

            var token = tokenResponse.Token;

            // signin/verifyState activity name used here to send my profile card after successful sign in.
            if ((activity.Type == MessageActivityType) || (activity.Name == SignInActivityName))
            {
                var command = ((string)stepContext.Values["command"]).ToUpperInvariant().Trim();

                switch (command)
                {
                    case Constants.MyProfile:
                        this.logger.LogInformation("My profile command triggered", new Dictionary<string, string>() { { "User", activity.From.Id }, { "AADObjectId", activity.From.AadObjectId } });
                        await this.MyProfileAsync(token, stepContext, cancellationToken).ConfigureAwait(false);
                        break;
                    case Constants.Search:
                        this.logger.LogInformation("Search command triggered", new Dictionary<string, string>() { { "User", activity.From.Id }, { "AADObjectId", activity.From.AadObjectId } });
                        await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(SearchCard.GetSearchCard())).ConfigureAwait(false);
                        break;
                    default:
                        await this.EditProfileAsync(token, stepContext, cancellationToken).ConfigureAwait(false);
                        break;
                }
            }

            // submit-invoke request at edit profile
            else if (activity.Type == InvokeActivityType)
            {
                await this.EditProfileAsync(token, stepContext, cancellationToken).ConfigureAwait(false);
            }

            return await stepContext.EndDialogAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Show profile details to user.
        /// </summary>
        /// <param name="token">User access token.</param>
        /// <param name="stepContext">Provides context for a step in a bot dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Tracking task.</returns>
        private async Task MyProfileAsync(string token, WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            try
            {
                var activity = stepContext.Context.Activity;
                IMessageActivity myProfileCardActivity;

                var userProfileDetails = await this.graphApiHelper.GetUserProfileAsync(token).ConfigureAwait(false);
                var userProfileCardId = Guid.NewGuid().ToString();

                if (userProfileDetails != null)
                {
                    this.logger.LogInformation($"User profile obtained from Graph for: {activity.Conversation.Id}.");
                    myProfileCardActivity = MessageFactory.Attachment(MyProfileCard.GetMyProfileCard(userProfileDetails, userProfileCardId));
                }
                else
                {
                    this.logger.LogInformation($"User profile obtained from Graph API is null for: {activity.Conversation.Id}.");
                    myProfileCardActivity = MessageFactory.Attachment(MyProfileCard.GetEmptyUserProfileCard(userProfileCardId));
                }

                var myProfileCardActivityResponse = await stepContext.Context.SendActivityAsync(myProfileCardActivity, cancellationToken).ConfigureAwait(false);
                await this.StoreUserProfileCardActivityInfoAsync(myProfileCardActivityResponse.Id, userProfileCardId, stepContext.Context).ConfigureAwait(false);

                this.logger.LogInformation("Profile sent to user", new Dictionary<string, string>() { { "User", activity.From.Id }, { "AADObjectId", activity.From.AadObjectId } });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error occured while executing My profile:  {stepContext.Context.Activity.Conversation.Id}.");
            }
        }

        /// <summary>
        /// Handle logic for edit profile task module.
        /// </summary>
        /// <param name="token">User access token.</param>
        /// <param name="stepContext">Provides context for a step in a bot dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Tracking task.</returns>
        private async Task EditProfileAsync(string token, WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            try
            {
                var activity = stepContext.Context.Activity;

                var userProfileDetails = new UserProfileDetailBase();
                var userProfileRequestData = JsonConvert.DeserializeObject<EditProfileCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase).ToString());

                userProfileDetails.AboutMe = userProfileRequestData.AboutMe;

                userProfileDetails.Skills = new List<string>();
                if (!string.IsNullOrEmpty(userProfileRequestData.Skills))
                {
                    var skills = userProfileRequestData.Skills.Split(';').Where(skillValue => !string.IsNullOrEmpty(skillValue));
                    userProfileDetails.Skills.AddRange(skills);
                }

                userProfileDetails.Interests = new List<string>();
                if (!string.IsNullOrEmpty(userProfileRequestData.Interests))
                {
                    var interests = userProfileRequestData.Interests.Split(';').Where(interestValue => !string.IsNullOrEmpty(interestValue));
                    userProfileDetails.Interests.AddRange(interests);
                }

                userProfileDetails.Schools = new List<string>();
                if (!string.IsNullOrEmpty(userProfileRequestData.Schools))
                {
                    var schools = userProfileRequestData.Schools.Split(';').Where(schoolValue => !string.IsNullOrEmpty(schoolValue));
                    userProfileDetails.Schools.AddRange(schools);
                }

                string userProfileDetailsData = JsonConvert.SerializeObject(userProfileDetails);
                bool isUserProfileUpdated = await this.graphApiHelper.UpdateUserProfileDetailsAsync(token, userProfileDetailsData).ConfigureAwait(false);

                if (!isUserProfileUpdated)
                {
                    await stepContext.Context.SendActivityAsync(Strings.FailedToUpdateProfile).ConfigureAwait(false);
                    this.logger.LogInformation($"Failure in saving data from task module to API for: {activity.Conversation.Id}.");
                }

                this.logger.LogInformation($"User profile updated using Graph API for conversation id: {activity.Conversation.Id}.");

                var userProfileCardId = ((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)["MyProfileCardId"].ToString();
                var userDetailsfromApi = await this.graphApiHelper.GetUserProfileAsync(token).ConfigureAwait(false);
                var userProfile = await this.storageHelper.GetUserProfileConversationDataAsync(userProfileCardId).ConfigureAwait(false);

                var updateProfileActivity = MessageFactory.Attachment(MyProfileCard.GetMyProfileCard(userDetailsfromApi, userProfileCardId));
                updateProfileActivity.Id = userProfile.MyProfileCardActivityId;
                updateProfileActivity.Conversation = activity.Conversation;
                await stepContext.Context.UpdateActivityAsync(updateProfileActivity, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error occured while posting my profile data to API for: {stepContext.Context.Activity.Conversation.Id}.");
                await stepContext.Context.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Save user profile activity information to Azure Table Storage, which is use to uniquely identify activity based on card id.
        /// </summary>
        /// <param name="myProfileCardActivityId">User profile card activity id.</param>
        /// <param name="myProfileCardId">Custom unique user profile card id.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>Tracking task.</returns>
        private async Task StoreUserProfileCardActivityInfoAsync(string myProfileCardActivityId, string myProfileCardId, ITurnContext turnContext)
        {
            string conversationId = turnContext.Activity.Conversation.Id;
            try
            {
                UserProfileActivityInfo userProfileActivityEntity = new UserProfileActivityInfo
                {
                    MyProfileCardActivityId = myProfileCardActivityId,
                    MyProfileCardId = myProfileCardId,
                };

                var isUserActivityInfoSaved = await this.storageHelper.UpsertUserProfileConversationDataAsync(userProfileActivityEntity).ConfigureAwait(false);
                if (!isUserActivityInfoSaved)
                {
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                    this.logger.LogInformation($"Saving data to table storage failed for: {conversationId}.");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Saving data to table storage failed for: {conversationId}.");
            }
        }
    }
}
