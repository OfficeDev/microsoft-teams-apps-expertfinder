// <copyright file="MessagingExtensionUserProfileCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Cards
{
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.ExpertFinder.Models.SharePoint;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Class having method to return messaging extension user profile details attachments.
    /// </summary>
    public static class MessagingExtensionUserProfileCard
    {
        /// <summary>
        /// Message extension command id for skills.
        /// </summary>
        private const string SkillsCommandId = "skills";

        /// <summary>
        /// Message extension command id for interests.
        /// </summary>
        private const string InterestCommandId = "interests";

        /// <summary>
        /// Message extension command id for schools.
        /// </summary>
        private const string SchoolsCommandId = "schools";

        /// <summary>
        /// Get user profile details messaging extension attachments for given user profiles and messaging extension command.
        /// </summary>
        /// <param name="userProfiles">Collection of user profile details.</param>
        /// <param name="commandId">Messaging extension command name.</param>
        /// <returns>List of user details messaging extension attachment.</returns>
        public static List<MessagingExtensionAttachment> GetUserDetailsCards(IList<UserProfileDetail> userProfiles, string commandId)
        {
            var messagingExtensionAttachments = new List<MessagingExtensionAttachment>();
            var cardContent = string.Empty;

            foreach (var userProfile in userProfiles)
            {
                switch (commandId)
                {
                    case SkillsCommandId:
                        cardContent = userProfile.Skills;
                        break;

                    case InterestCommandId:
                        cardContent = userProfile.Interests;
                        break;

                    case SchoolsCommandId:
                        cardContent = userProfile.Schools;
                        break;
                }

                var userCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = userProfile.PreferredName,
                            Weight = AdaptiveTextWeight.Bolder,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = userProfile.JobTitle,
                            Wrap = true,
                            Spacing = AdaptiveSpacing.None,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = Strings.AboutMeTitle,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = userProfile.AboutMe,
                            IsSubtle = true,
                            Wrap = true,
                            Spacing = AdaptiveSpacing.None,
                        },
                    },
                };
                ThumbnailCard previewCard = new ThumbnailCard
                {
                    Title = $"<strong>{userProfile.PreferredName}</strong>",
                    Subtitle = userProfile.JobTitle,
                    Text = cardContent,
                };
                messagingExtensionAttachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = userCard,
                }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
            }

            return messagingExtensionAttachments;
        }
    }
}