// <copyright file="MyProfileCard.cs" company="Microsoft">
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
    /// Class having methods related to user profile attachment.
    /// </summary>
    public static class MyProfileCard
    {
        /// <summary>
        /// Base uri to view user profile.
        /// </summary>
        private const string GoToProfileUrl = "https://delve.office.com/";

        /// <summary>
        /// text that triggers go to profile action.
        /// </summary>
        private const string GoToProfileCommand = "Go to profile";

        /// <summary>
        /// Get the user profile card attachment for given user profile.
        /// </summary>
        /// <param name="userProfile">User profile details.</param>
        /// <param name="profileCardId">User profile unique activity card Id.</param>
        /// <returns>User profile details card attachment for given user.</returns>
        public static Attachment GetMyProfileCard(UserProfileDetail userProfile, string profileCardId)
        {
            var skills = userProfile.Skills.Count > 0 ? string.Join(";", userProfile.Skills) : Strings.NoneText;
            var interests = userProfile.Interests.Count > 0 ? string.Join(";", userProfile.Interests) : Strings.NoneText;
            var schools = userProfile.Schools.Count > 0 ? string.Join(";", userProfile.Schools) : Strings.NoneText;

            var myProfileCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = userProfile.DisplayName,
                        Wrap = true,
                        Weight = AdaptiveTextWeight.Bolder,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = userProfile.JobTitle,
                        Wrap = true,
                        IsSubtle = true,
                        Weight = AdaptiveTextWeight.Default,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = userProfile.AboutMe,
                        Wrap = true,
                        Spacing = AdaptiveSpacing.Small,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.EditProfileTitle,
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                 Type = Constants.FetchActionType,
                            },
                            Command = Constants.MyProfile,
                            MyProfileCardId = profileCardId,
                        },
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
                                   Text = Strings.SkillsTitle,
                                   Separator = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                                },
                                new AdaptiveTextBlock
                                {
                                   Text = skills,
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
                                    Title = GoToProfileCommand,
                                    Url = new Uri($"{GoToProfileUrl}?u={userProfile.Id}&v=profiledetails"),
                                },
                            },
                        },
                    },
                },
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = myProfileCard,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        ///  Get the user profile edit card for given user profile.
        /// </summary>
        /// <param name="userProfile">User profile details.</param>
        /// <param name="profileCardId">User profile unique card Id.</param>
        /// <param name="appBaseUrl">Applicaion base uri.</param>
        /// <returns>User profile details edit card attachment for given user.</returns>
        public static Attachment GetEditProfileCard(UserProfileDetail userProfile, string profileCardId, string appBaseUrl)
        {
            var skills = userProfile.Skills.Count > 0 ? string.Join(";", userProfile.Skills) : Strings.NoneText;
            var interests = userProfile.Interests.Count > 0 ? string.Join(";", userProfile.Interests) : Strings.NoneText;
            var schools = userProfile.Schools.Count > 0 ? string.Join(";", userProfile.Schools) : Strings.NoneText;

            var myProfileCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Text = Strings.FullNameTitle,
                                Size = AdaptiveTextSize.Small,
                                Wrap = true,
                            },
                            new AdaptiveTextBlock
                            {
                                Text = userProfile.DisplayName,
                                Wrap = true,
                                Spacing = AdaptiveSpacing.None,
                            },

                            new AdaptiveTextBlock
                            {
                                Text = Strings.AboutMeTitle,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextInput
                            {
                                Placeholder = Strings.AboutMePlaceHolderText,
                                IsMultiline = true,
                                Style = AdaptiveTextInputStyle.Text,
                                Id = "aboutme",
                                MaxLength = 300,
                                Value = userProfile.AboutMe,
                                Spacing = AdaptiveSpacing.None,
                            },
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
                                                Url = new Uri($"{appBaseUrl}/Artifacts/validationIcon.png"),
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
                                                Text = Strings.ValidationTaskModuleMessage,
                                                Wrap = true,
                                            },
                                        },
                                    },
                                },
                            },

                            new AdaptiveTextBlock
                            {
                                Text = Strings.InterestTitle,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextInput
                            {
                                Placeholder = Strings.InterestsPlaceHolderText,
                                IsMultiline = true,
                                Style = AdaptiveTextInputStyle.Text,
                                Id = "interests",
                                MaxLength = 100,
                                Value = interests,
                                Spacing = AdaptiveSpacing.None,
                            },
                            new AdaptiveTextBlock
                            {
                                Text = Strings.SchoolsTitle,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextInput
                            {
                                Placeholder = Strings.SchoolsPlaceHolderText,
                                IsMultiline = true,
                                Style = AdaptiveTextInputStyle.Text,
                                Id = "schools",
                                MaxLength = 200,
                                Value = schools,
                                Spacing = AdaptiveSpacing.None,
                            },
                            new AdaptiveTextBlock
                            {
                                Text = Strings.SkillsTitle,
                                Wrap = true,
                                Size = AdaptiveTextSize.Small,
                            },
                            new AdaptiveTextInput
                            {
                                Placeholder = Strings.SkillsPlaceHolderText,
                                IsMultiline = true,
                                Style = AdaptiveTextInputStyle.Text,
                                Id = "skills",
                                MaxLength = 100,
                                Value = skills,
                                Spacing = AdaptiveSpacing.None,
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction()
                    {
                        Title = Strings.UpdateTitle,
                        Data = new AdaptiveCardAction
                        {
                            Command = Constants.MyProfile,
                            MyProfileCardId = profileCardId,
                        },
                    },
                },
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = myProfileCard,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Get the user profile card if user profile details not present for user.
        /// </summary>
        /// <param name="profileCardId">User profile unique card id.</param>
        /// <returns>User profile details card attachment.</returns>
        public static Attachment GetEmptyUserProfileCard(string profileCardId)
        {
            var emptyProfileCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.EmptyProfileCardContent,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.EditProfileTitle,
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                 Type = Constants.FetchActionType,
                            },
                            Command = Constants.MyProfile,
                            MyProfileCardId = profileCardId,
                        },
                    },
                },
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = emptyProfileCard,
            };
            return adaptiveCardAttachment;
        }
    }
}
