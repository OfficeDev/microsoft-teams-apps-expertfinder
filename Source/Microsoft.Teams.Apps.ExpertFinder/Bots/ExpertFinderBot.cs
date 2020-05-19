// <copyright file="ExpertFinderBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Bots
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Rest;
    using Microsoft.Teams.Apps.ExpertFinder.Cards;
    using Microsoft.Teams.Apps.ExpertFinder.Common;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Class that handles the teams activity of Expert Finder bot and messaging extension.
    /// </summary>
    public class ExpertFinderBot : TeamsActivityHandler
    {
        /// <summary>
        /// Messaging extension result type.
        /// </summary>
        private const string MessagingExtenstionResultType = "result";

        /// <summary>
        /// Messaging extension message type.
        /// </summary>
        private const string MessagingExtensionMessageType = "message";

        /// <summary>
        /// Messaging extension auth type.
        /// </summary>
        private const string MessagingExtensionAuthType = "auth";

        /// <summary>
        /// Microsoft Graph resource URI.
        /// </summary>
        private const string GraphResourceUri = "https://graph.microsoft.com";

        /// <summary>
        /// Messaging extension default parameter value.
        /// </summary>
        private const string MessagingExtensionInitialParameterName = "initialRun";

        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int TaskModuleHeight = 600;

        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int TaskModuleWidth = 600;

        // Async retry policy that will wait and retry as many times as there are provided sleep durations which is
        // an exponentially backing-off, jittered manner, making sure to mitigate any correlations.
        private static readonly AsyncRetryPolicy RetryPolicy = Policy.Handle<HttpOperationException>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 5));

        /// <summary>
        /// State management object for maintaining conversation state.
        /// </summary>
        private readonly BotState conversationState;

        /// <summary>
        /// Base class for all bot dialogs.
        /// </summary>
        private readonly Dialog rootDialog;

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// Helper for working with Microsoft Graph API.
        /// </summary>
        private readonly IGraphApiHelper graphApiHelper;

        /// <summary>
        /// Helper for acquiring AAD token for given resource.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Helper for custom JWT token generation, token validation and acquiring token for given resource.
        /// </summary>
        private readonly ICustomTokenHelper customTokenHelper;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Helper object for working with SharePoint REST API.
        /// </summary>
        private readonly ISharePointApiHelper sharePointApiHelper;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Expert Finder bot.
        /// </summary>
        private readonly BotSettings botSettings;

        private readonly IStatePropertyAccessor<DialogState> dialogStatePropertyAccessor;
        private readonly IStatePropertyAccessor<UserData> userDataPropertyAccessor;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpertFinderBot"/> class.
        /// </summary>
        /// <param name="conversationState">State management object for maintaining conversation state.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        /// <param name="rootDialog">Root dialog.</param>
        /// <param name="graphApiHelper">Helper for working with Microsoft Graph API.</param>
        /// <param name="tokenHelper">Helper for JWT token generation and validation.</param>
        /// <param name="sharePointApiHelper">Helper object for working with SharePoint REST API.</param>
        /// <param name="botSettings">A set of key/value application configuration properties for Expert Finder bot.</param>
        /// <param name="customTokenHelper">Helper for AAD token generation.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ExpertFinderBot(ConversationState conversationState, UserState userState, MainDialog rootDialog, IGraphApiHelper graphApiHelper, ITokenHelper tokenHelper, ISharePointApiHelper sharePointApiHelper, ICustomTokenHelper customTokenHelper, IOptionsMonitor<BotSettings> botSettings, ILogger<ExpertFinderBot> logger)
        {
            this.conversationState = conversationState;
            this.userState = userState;
            this.rootDialog = rootDialog;
            this.graphApiHelper = graphApiHelper;
            this.tokenHelper = tokenHelper;
            this.sharePointApiHelper = sharePointApiHelper;
            this.botSettings = botSettings.CurrentValue;
            this.customTokenHelper = customTokenHelper;
            this.logger = logger;

            this.dialogStatePropertyAccessor = this.conversationState.CreateProperty<DialogState>(nameof(DialogState));
            this.userDataPropertyAccessor = this.conversationState.CreateProperty<UserData>("ConversationData");    // For compatibility with v1 of Expert Finder
        }

        /// <summary>
        /// Method will be invoked on each bot turn.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            if (!this.IsActivityFromExpectedTenant(turnContext))
            {
                this.logger.LogInformation($"Unexpected tenant id {turnContext.Activity.Conversation.TenantId}", SeverityLevel.Warning);
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.InvalidTenant)).ConfigureAwait(false);
            }
            else
            {
                // Get the current culture info to use in resource files
                string locale = turnContext.Activity.Entities?.Where(t => t.Type == "clientInfo").First().Properties["locale"].ToString();
                if (!string.IsNullOrEmpty(locale))
                {
                    CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(locale);
                }

                await base.OnTurnAsync(turnContext, cancellationToken).ConfigureAwait(false);

                // Save any state changes that might have occured during the turn.
                await this.conversationState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
                await this.userState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Method that checks teams signin verify state, check if token exists.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnTeamsSigninVerifyStateAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            await this.rootDialog.RunAsync(turnContext, this.dialogStatePropertyAccessor, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// When OnTurn method receives a message activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var activity = turnContext.Activity;
                await this.SendTypingIndicatorAsync(turnContext).ConfigureAwait(false);
                await this.rootDialog.RunAsync(turnContext, this.dialogStatePropertyAccessor, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in message activity of bot for {turnContext.Activity.Conversation.Id}");
                throw;
            }
        }

        /// <summary>
        /// Implemented this to provide logic when bot is added, to implement bot's welcome logic.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Welcome card  when bot is added first time by user.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {membersAdded.Count}");

            if (membersAdded.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                var userData = await this.userDataPropertyAccessor.GetAsync(turnContext, () => new UserData()).ConfigureAwait(false);
                if (userData?.IsWelcomeCardSent == null || userData?.IsWelcomeCardSent == false)
                {
                    var userWelcomeCardAttachment = WelcomeCard.GetCard(this.botSettings.AppBaseUri);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment)).ConfigureAwait(false);

                    userData.IsWelcomeCardSent = true;
                }
            }
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="taskModuleRequestData">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequestData, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            var userSearchTaskModuleDetails = JsonConvert.DeserializeObject<EditProfileCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase).ToString());
            string command = userSearchTaskModuleDetails.Command;

            try
            {
                var userGraphAccessToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id, GraphResourceUri).ConfigureAwait(false);

                if (userGraphAccessToken == null)
                {
                    await turnContext.SendActivityAsync(Strings.NotLoggedInText).ConfigureAwait(false);
                    await this.rootDialog.RunAsync(turnContext, this.dialogStatePropertyAccessor, cancellationToken).ConfigureAwait(false);
                    return null;
                }
                else
                {
                    switch (command)
                    {
                        case Constants.Search:
                            this.logger.LogInformation("Search fetch activity called");
                            var apiAuthToken = this.customTokenHelper.GenerateAPIAuthToken(aadObjectId: activity.From.AadObjectId, serviceURL: activity.ServiceUrl, fromId: activity.From.Id, jwtExpiryMinutes: 60);
                            return new TaskModuleResponse
                            {
                                Task = new TaskModuleContinueResponse
                                {
                                    Value = new TaskModuleTaskInfo()
                                    {
                                        Url = $"{this.botSettings.AppBaseUri}/?token={apiAuthToken}&telemetry={this.botSettings.AppInsightsInstrumentationKey}&theme={{theme}}",
                                        Height = TaskModuleHeight,
                                        Width = TaskModuleWidth,
                                        Title = Strings.SearchTaskModuleTitle,
                                    },
                                },
                            };

                        case Constants.MyProfile:
                            this.logger.LogInformation("My profile fetch activity called");
                            var userProfileDetails = await this.graphApiHelper.GetUserProfileAsync(userGraphAccessToken).ConfigureAwait(false);
                            if (userProfileDetails == null)
                            {
                                this.logger.LogInformation("User profile details obtained from Graph API is null.");
                                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                                return null;
                            }
                            else
                            {
                                return new TaskModuleResponse
                                {
                                    Task = new TaskModuleContinueResponse
                                    {
                                        Value = new TaskModuleTaskInfo()
                                        {
                                            Card = MyProfileCard.GetEditProfileCard(userProfileDetails, userSearchTaskModuleDetails.MyProfileCardId, this.botSettings.AppBaseUri),
                                            Height = TaskModuleHeight,
                                            Width = TaskModuleWidth,
                                            Title = Strings.EditProfileTitle,
                                        },
                                    },
                                };
                            }

                        default:
                            this.logger.LogInformation($"Invalid command for task module fetch activity.Command is : {command} ");
                            await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                            return null;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in fetch action of task module.");
                return null;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="taskModuleRequestData">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequestData, CancellationToken cancellationToken)
        {
            var valuesFromTaskModule = JsonConvert.DeserializeObject<SearchSubmitAction>(taskModuleRequestData.Data?.ToString());
            try
            {
                if (valuesFromTaskModule == null)
                {
                    this.logger.LogInformation($"Request data obtained on task module submit action is null.");
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                    return null;
                }

                switch (valuesFromTaskModule.Command.ToUpperInvariant().Trim())
                {
                    case Constants.MyProfile:
                        this.logger.LogInformation("Activity type is invoke submit from my profile command");
                        await this.rootDialog.RunAsync(turnContext, this.dialogStatePropertyAccessor, cancellationToken).ConfigureAwait(false);
                        break;

                    case Constants.Search:
                        this.logger.LogInformation("Activity type is invoke submit from search command");

                        foreach (var profile in valuesFromTaskModule.UserProfiles)
                        {
                            // Bot is expected to send multiple user profile cards which may cross the threshold limit of bot messages/sec, hence adding the retry logic.
                            await RetryPolicy.ExecuteAsync(async () =>
                            {
                                await turnContext.SendActivityAsync(MessageFactory.Attachment(SearchCard.GetUserCard(profile)), cancellationToken).ConfigureAwait(false);
                            }).ConfigureAwait(false);
                        }

                        break;
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in submit action of task module.");
                return null;
            }
        }

        /// <summary>
        /// When OnTurn method receives a compose extension query invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="query">Messaging extension query request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents messaging extension response.</returns>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            try
            {
                var messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContext.Activity.Value.ToString());

                // Execute code when parameter name is initial run.
                if (messageExtensionQuery.Parameters.First().Name == MessagingExtensionInitialParameterName)
                {
                    this.logger.LogInformation("Executing initial run parameter from messaging extension.");

                    // Get access token for user.if already authenticated, we will get token.
                    // If user is not signed in, send sign in link in messaging extension.
                    var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, this.botSettings.OAuthConnectionName, messageExtensionQuery.State, cancellationToken).ConfigureAwait(false);

                    if (tokenResponse == null)
                    {
                        var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, this.botSettings.OAuthConnectionName, cancellationToken).ConfigureAwait(false);
                        return new MessagingExtensionResponse
                        {
                            ComposeExtension = new MessagingExtensionResult
                            {
                                Type = MessagingExtensionAuthType,
                                SuggestedActions = new MessagingExtensionSuggestedAction
                                {
                                    Actions = new List<CardAction>
                                    {
                                        new CardAction
                                        {
                                            Type = ActionTypes.OpenUrl,
                                            Value = signInLink,
                                            Title = Strings.SignInCardText,
                                        },
                                    },
                                },
                            },
                        };
                    }
                }

                return await this.HandleMessagingExtensionSearchQueryAsync(turnContext).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in handling invoke action from messaging extension.");
                return null;
            }
        }

        /// <summary>
        /// Handles messaging extension user search query request.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents messaging extension response containing user profile details received from SharePoint API based on user search query.</returns>
        private async Task<MessagingExtensionResponse> HandleMessagingExtensionSearchQueryAsync(ITurnContext<IInvokeActivity> turnContext)
        {
            try
            {
                var activity = turnContext.Activity;
                var messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(activity.Value.ToString());
                var searchQuery = messageExtensionQuery.Parameters.First().Value.ToString();

                this.logger.LogInformation($"searchQuery : {searchQuery} commandId : {messageExtensionQuery.CommandId}");

                // Get SharePoint user access token.
                var token = await this.tokenHelper.GetUserTokenAsync(activity.From.Id, this.botSettings.SharePointSiteUrl).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.logger.LogInformation($"Token not obtained while handling messaging extension query for {activity.Conversation.Id}.");
                    return null;
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, token).ConfigureAwait(false),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception while handling messaging extension query");
                return null;
            }
        }

        /// <summary>
        /// Get messaging extension search result based on user search query and command.
        /// </summary>
        /// <param name="searchQuery">User search query text.</param>
        /// <param name="commandId">Messaging extension command id e.g. skills, interests, schools.</param>
        /// <param name="token">User access token.</param>
        /// <returns>A task that represents compose extension result containing user profile details.</returns>
        private async Task<MessagingExtensionResult> GetSearchResultAsync(string searchQuery, string commandId, string token)
        {
            try
            {
                MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult();

                // On intial run searchQuery value is "true".
                if (searchQuery == "true")
                {
                    composeExtensionResult.Type = MessagingExtensionMessageType;
                    composeExtensionResult.Text = Strings.DefaultCardContentME;
                }
                else
                {
                    composeExtensionResult.Type = MessagingExtenstionResultType;
                    composeExtensionResult.AttachmentLayout = AttachmentLayoutTypes.List;
                    var userProfiles = await this.sharePointApiHelper.GetUserProfilesAsync(searchQuery, new List<string>() { commandId }, token, this.botSettings.SharePointSiteUrl).ConfigureAwait(false);

                    if (userProfiles.Count > 0)
                    {
                        composeExtensionResult.Attachments = MessagingExtensionUserProfileCard.GetUserDetailsCards(userProfiles, commandId);
                    }
                    else
                    {
                        this.logger.LogInformation("User profile obtained from sharepoint search service is null.");
                    }
                }

                return composeExtensionResult;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while generating result for messaging extension");
                return null;
            }
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents typing indicator activity.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
            catch (Exception ex)
            {
                // Do not fail on errors sending the typing indicator
                this.logger.LogWarning(ex, $"Failed to send a typing indicator: {ex.Message}");
            }
        }

        /// <summary>
        /// Verify if the tenant id in the message is the same tenant id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A boolean, true if tenant provided is expexted tenant.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId.Equals(this.botSettings.TenantId, StringComparison.OrdinalIgnoreCase);
        }
    }
}