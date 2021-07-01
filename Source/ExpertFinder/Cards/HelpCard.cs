// <copyright file="HelpCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.ExpertFinder.Common;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Class that contains method for help card attachment.
    /// </summary>
    public static class HelpCard
    {
        /// <summary>
        /// Get help card attchment that will give available commands to user if user has provided invalid command.
        /// </summary>
        /// <returns>Help adaptive card attachment.</returns>
        public static Attachment GetHelpCard()
        {
            AdaptiveCard helpCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.HelpMessage,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.SearchTitle,
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                DisplayText = Strings.SearchTitle,
                            },
                            Command = Constants.Search,
                        },
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.MyProfileTitle,
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                DisplayText = Strings.MyProfileTitle,
                            },
                            Command = Constants.MyProfile,
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = helpCard,
            };
        }
    }
}