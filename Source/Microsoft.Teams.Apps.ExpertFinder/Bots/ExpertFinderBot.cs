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
    /// <typeparam name="T">Generic class.</typeparam>
    public class ExpertFinderBot<T> : TeamsActivityHandler
        where T : Dialog
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
        /// Microsoft Graph api base uri.
        /// </summary>
        private const string GraphAPIBaseURL = "https://graph.microsoft.com";

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
        private readonly Dialog dialog;

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// Helper for working with Microsoft Graph api.
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
        /// Application base uri.
        /// </summary>
        private readonly string appBaseUrl;

        /// <summary>
        /// Helper object for working with SharePoint rest api.
        /// </summary>
        private readonly ISharePointApiHelper sharePointApiHelper;

        /// <summary>
        /// AADv1 bot connection name.
        /// </summary>
        private readonly string connectionName;

        /// <summary>
        /// SharePoint site Uri.
        /// </summary>
        private readonly string sharePointSiteUri;

        /// <summary>
        /// Application Insights instrumentation key which we passes to client application.
        /// </summary>
        private readonly string appInsightsInstrumentationKey;

        /// <summary>
        /// Tenant id.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Expert Finder bot.
        /// </summary>
        private readonly BotSettings options;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpertFinderBot{T}"/> class.
        /// </summary>
        /// <param name="conversationState">State management object for maintaining conversation state.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        /// <param name="dialog">Base class for bot dialog.</param>
        /// <param name="graphApiHelper">Helper for working with Microsoft Graph api.</param>
        /// <param name="tokenHelper">Helper for JWT token generation and validation.</param>
        /// <param name="sharePointApiHelper">Helper object for working with SharePoint rest api.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for Expert Finder bot.</param>
        /// <param name="customTokenHelper">Helper for AAD token generation.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ExpertFinderBot(ConversationState conversationState, UserState userState, T dialog, IGraphApiHelper graphApiHelper, ITokenHelper tokenHelper, ISharePointApiHelper sharePointApiHelper, ICustomTokenHelper customTokenHelper, IOptionsMonitor<BotSettings> optionsAccessor, ILogger<ExpertFinderBot<T>> logger)
        {
            this.conversationState = conversationState;
            this.userState = userState;
            this.dialog = dialog;
            this.graphApiHelper = graphApiHelper;
            this.tokenHelper = tokenHelper;
            this.sharePointApiHelper = sharePointApiHelper;
            this.options = optionsAccessor.CurrentValue;
            this.appBaseUrl = this.options.AppBaseUri;
            this.connectionName = this.options.ConnectionName;
            this.sharePointSiteUri = this.options.SharePointSiteUrl;
            this.appInsightsInstrumentationKey = this.options.AppInsightsInstrumentationKey;
            this.tenantId = this.options.TenantId;
            this.customTokenHelper = customTokenHelper;
            this.logger = logger;
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
                this.logger.LogInformation($"Unexpected tenant Id {turnContext.Activity.Conversation.TenantId}", SeverityLevel.Warning);
                await turnContext.SendActivityAsync(activity: MessageFactory.Text(Strings.InvalidTenant)).ConfigureAwait(false);
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
                await this.conversationState.SaveChangesAsync(turnContext: turnContext, force: false, cancellationToken: cancellationToken).ConfigureAwait(false);
                await this.userState.SaveChangesAsync(turnContext: turnContext, force: false, cancellationToken: cancellationToken).ConfigureAwait(false);
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
            await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
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

                // if card from messaging extension is sent to the bot conversation.
                if (activity.Attachments != null)
                {
                    return;
                }
                else
                {
                    var command = activity.Text;
                    await this.SendTypingIndicatorAsync(turnContext).ConfigureAwait(false);
                    if (activity.Text == null && activity.Value != null && activity.Type == ActivityTypes.Message)
                    {
                        command = JToken.Parse(activity.Value.ToString()).SelectToken("command").ToString();
                    }

                    switch (command.ToUpperInvariant().Trim())
                    {
                        case Constants.MyProfile:
                            await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                            break;
                        case Constants.Search:
                            await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                            break;
                        default:
                            await turnContext.SendActivityAsync(MessageFactory.Attachment(HelpCard.GetHelpCard())).ConfigureAwait(false);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in message activity of bot for {turnContext.Activity.Conversation.Id}");
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
            this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                var userStateAccessors = this.userState.CreateProperty<ConversationData>(nameof(ConversationData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new ConversationData()).ConfigureAwait(false);
                if (userdata?.IsWelcomeCardSent == null || userdata?.IsWelcomeCardSent == false)
                {
                    userdata.IsWelcomeCardSent = true;
                    var userWelcomeCardAttachment = WelcomeCard.GetCard(this.appBaseUrl);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment)).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="membersRemoved">List of members removed.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext?.Activity;

            this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");
            if (activity.MembersAdded.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                var userStateAccessors = this.userState.CreateProperty<ConversationData>(nameof(ConversationData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new ConversationData()).ConfigureAwait(false);
                userdata.IsWelcomeCardSent = false;
                await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
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
                var userGraphAccessToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id, GraphAPIBaseURL).ConfigureAwait(false);

                if (userGraphAccessToken == null)
                {
                    await turnContext.SendActivityAsync(Strings.NotLoginText).ConfigureAwait(false);
                    await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                    return default;
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
                                        Url = $"{this.appBaseUrl}/?token={apiAuthToken}&telemetry={this.appInsightsInstrumentationKey}&theme=" + "{theme}",
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
                                this.logger.LogInformation("UserProfile details obtained from graph api is null.");
                                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                                return default;
                            }
                            else
                            {
                                return new TaskModuleResponse
                                {
                                    Task = new TaskModuleContinueResponse
                                    {
                                        Value = new TaskModuleTaskInfo()
                                        {
                                            Card = MyProfileCard.GetEditProfileCard(userProfileDetails, userSearchTaskModuleDetails.MyProfileCardId, this.appBaseUrl),
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
                            return default;
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in fetch action of task module.");
                return default;
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
                    return default;
                }

                switch (valuesFromTaskModule.Command.ToUpperInvariant().Trim())
                {
                    case Constants.MyProfile:
                        this.logger.LogInformation("Activity type is invoke submit from my profile command");
                        await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                        break;
                    case Constants.Search:
                        this.logger.LogInformation("Activity type is invoke submit from search command");
                        List<IActivity> selectedUserActivities = new List<IActivity>();
                        valuesFromTaskModule.UserProfiles.ForEach(userProfile => selectedUserActivities.Add(MessageFactory.Attachment(SearchCard.GetUserCard(userProfile))));

                        // Bot is expected to send multiple user profile cards which may cross the threshold limit of bot messages/sec, hence adding the retry logic.
                        await RetryPolicy.ExecuteAsync(async () =>
                        {
                            await turnContext.SendActivitiesAsync(selectedUserActivities.ToArray(), cancellationToken).ConfigureAwait(false);
                        }).ConfigureAwait(false);
                        break;
                }

                return default;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in submit action of task module.");
                return default;
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
                    var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, this.connectionName, messageExtensionQuery.State, cancellationToken).ConfigureAwait(false);

                    if (tokenResponse == null)
                    {
                        var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, this.connectionName, cancellationToken).ConfigureAwait(false);
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
                                            Title = Strings.SigninCardText,
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
            }

            return default;
        }

        /// <summary>
        /// Handles messaging extension user search query request.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents messaging extension response containing user profile details received from SharePoint api based on user search query.</returns>
        private async Task<MessagingExtensionResponse> HandleMessagingExtensionSearchQueryAsync(ITurnContext<IInvokeActivity> turnContext)
        {
            try
            {
                var activity = turnContext.Activity;
                var messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(activity.Value.ToString());
                var searchQuery = messageExtensionQuery.Parameters.First().Value.ToString();

                this.logger.LogInformation($"searchQuery : {searchQuery} commandId : {messageExtensionQuery.CommandId}");

                // Get SharePoint user access token.
                var token = await this.tokenHelper.GetUserTokenAsync(activity.From.Id, this.sharePointSiteUri).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.logger.LogInformation($"Token not obtained while handling messaging extension query for {activity.Conversation.Id}.");
                    return default;
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, token).ConfigureAwait(false),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception while handling messaging extension query");
                return default;
            }
        }

        /// <summary>
        /// Get messaging extension search result based on user search query and command.
        /// </summary>
        /// <param name="searchQuery">User search query text.</param>
        /// <param name="commandId">Messaging extension command id e.g. skills, interests, schools.</param>
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
                    var userProfiles = await this.sharePointApiHelper.GetUserProfilesAsync(searchQuery, new List<string>() { commandId }, token, this.sharePointSiteUri).ConfigureAwait(false);

                    if (userProfiles.Count > 0)
                    {
                        composeExtensionResult.Attachments = MessagingExtensionUserProfileCard.GetUserDetailsCards(userProfiles, commandId);
                    }
                    else
                    {
                        this.logger.LogInformation("UserProfile obtained from sharepoint search service is null.");
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
            var typingActivity = turnContext.Activity.CreateReply(locale: CultureInfo.CurrentCulture.Name);
            typingActivity.Type = ActivityTypes.Typing;
            await turnContext.SendActivityAsync(typingActivity).ConfigureAwait(false);
        }

        /// <summary>
        /// Verify if the tenant id in the message is the same tenant id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A boolean, true if tenant provided is expexted tenant.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId.Equals(this.tenantId, StringComparison.OrdinalIgnoreCase);
        }
    }
}