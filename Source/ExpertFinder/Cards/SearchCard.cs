// <copyright file="SearchCard.cs" company="Microsoft">
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
    /// Class having methods related to user search card attachment and user profile details card attachment.
    /// </summary>
    public static class SearchCard
    {
        /// <summary>
        /// Url to initiate teams 1:1 chat with user.
        /// </summary>
        private const string InitiateChatUrl = "https://teams.microsoft.com/l/chat/0/0?users=";

        /// <summary>
        /// Card attachment to show on search command.
        /// </summary>
        /// <returns>Fetch action user search card attachment.</returns>
        public static Attachment GetSearchCard()
        {
            var searchCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.SearchCardContent,
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
                                 Type = Constants.FetchActionType,
                            },
                            Command = Constants.Search,
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = searchCard,
            };
        }

        /// <summary>
        /// User detail card attachment for given user profile.
        /// </summary>
        /// <param name="userDetail">User profile details.</param>
        /// <returns>User profile details card attachment.</returns>
        public static Attachment GetUserCard(Models.SharePoint.UserProfileDetail userDetail)
        {
            var skills = string.IsNullOrEmpty(userDetail.Skills) ? Strings.NoneText : userDetail.Skills;
            var interests = string.IsNullOrEmpty(userDetail.Interests) ? Strings.NoneText : userDetail.Interests;
            var schools = string.IsNullOrEmpty(userDetail.Schools) ? Strings.NoneText : userDetail.Schools;

            var userDetailCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = userDetail.PreferredName,
                        Wrap = true,
                        Weight = AdaptiveTextWeight.Bolder,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = userDetail.JobTitle,
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = $"_{Strings.SkillsTitle}_",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = skills,
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveOpenUrlAction
                    {
                        Title = Strings.ChatTitle,
                        Url = new System.Uri($"{InitiateChatUrl}{userDetail.WorkEmail}"),
                    },
                    new AdaptiveShowCardAction
                    {
                        Title = Strings.DetailsTitle,
                        Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
                        {
                            Body = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                   Text = Strings.AboutMeTitle,
                                   Separator = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = userDetail.AboutMe,
                                   Wrap = true,
                                   Spacing = AdaptiveSpacing.None,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = Strings.InterestTitle,
                                   Separator = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = interests,
                                   Wrap = true,
                                   Spacing = AdaptiveSpacing.None,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = Strings.SchoolsTitle,
                                   Separator = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = schools,
                                   Wrap = true,
                                   Spacing = AdaptiveSpacing.None,
                                },
                            },
                            Actions = new List<AdaptiveAction>
                            {
                                new AdaptiveOpenUrlAction
                                {
                                    Title = Strings.GotoProfileTitle,
                                    Url = new System.Uri($"{userDetail.Path}&v=profiledetails"),
                                },
                            },
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = userDetailCard,
            };
        }
    }
}