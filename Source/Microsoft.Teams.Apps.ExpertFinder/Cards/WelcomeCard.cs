// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.ExpertFinder.Common;
    using Microsoft.Teams.Apps.ExpertFinder.Models;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Implements Welcome Card.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// This method will construct the user welcome card when bot is added by user.
        /// </summary>
        /// <param name="appBaseUrl">Application base url.</param>
        /// <returns>User welcome card attchment.</returns>
        public static Attachment GetCard(string appBaseUrl)
        {
            var userWelcomeCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{appBaseUrl}/Artifacts/appLogo.png"),
                                        Size = AdaptiveImageSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Large,
                                        Wrap = true,
                                        Text = Strings.WelcomeText,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Default,
                                        Wrap = true,
                                        Text = Strings.WelcomeCardContent,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"**{Strings.SearchTitle}**: {Strings.SearchWelcomeCardContent}",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"**{Strings.MyProfileTitle}**: {Strings.MyProfileWelcomeCardContent}",
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
                Content = userWelcomeCard,
            };
        }
    }
}
