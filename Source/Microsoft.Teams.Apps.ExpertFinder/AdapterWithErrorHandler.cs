// <copyright file="AdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder
{
    using System;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Log any leaked exception from the application.
    /// </summary>
    public class AdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Object that passes the application configuration key-values.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="conversationState">State management object for maintaining conversation state.</param>
        public AdapterWithErrorHandler(IConfiguration configuration, ILogger<IBotFrameworkHttpAdapter> logger, ConversationState conversationState)
            : base(configuration)
        {
            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError(exception, $"Exception caught : {exception.Message}");

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);

                if (conversationState != null)
                {
                    logger.LogTrace($"Clearing conversation state for {turnContext.Activity?.Conversation?.Id}");
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"Exception caught on attempting to Delete ConversationState : {ex.Message}");
                    }
                }
            };
        }
    }
}
