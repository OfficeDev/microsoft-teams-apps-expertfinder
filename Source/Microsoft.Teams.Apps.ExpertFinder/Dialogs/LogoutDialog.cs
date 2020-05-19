// <copyright file="LogoutDialog.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Dialogs
{
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.ExpertFinder.Resources;

    /// <summary>
    /// Dialog for handling interruption.
    /// </summary>
    public class LogoutDialog : ComponentDialog
    {
        /// <summary>
        /// Text that triggers logout action.
        /// </summary>
        private static readonly ISet<string> LogoutCommands = new HashSet<string> { "LOGOUT", "SIGNOUT", "LOG OUT", "SIGN OUT" };

        /// <summary>
        /// Bot OAuth connection name.
        /// </summary>
        private readonly string connectionName;

        /// <summary>
        /// Initializes a new instance of the <see cref="LogoutDialog"/> class.
        /// </summary>
        /// <param name="id">Dialog Id.</param>
        /// <param name="connectionName">AADv1 connection name.</param>
        public LogoutDialog(string id, string connectionName)
            : base(id)
        {
            this.connectionName = connectionName;
        }

        /// <summary>
        /// Called when the dialog is started and pushed onto the parent's dialog stack.
        /// </summary>
        /// <param name="dialogContext">The inner Microsoft.Bot.Builder.Dialogs.DialogContext for the current turn of conversation.</param>
        /// <param name="options">Optional, initial information to pass to the dialog.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        protected override async Task<DialogTurnResult> OnBeginDialogAsync(DialogContext dialogContext, object options, CancellationToken cancellationToken = default)
        {
            var result = await this.InterruptAsync(dialogContext, cancellationToken).ConfigureAwait(false);
            if (result != null)
            {
                return result;
            }

            return await base.OnBeginDialogAsync(dialogContext, options, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Called when the dialog is _continued_, where it is the active dialog and the user replies with a new activity.
        /// </summary>
        /// <param name="dialogContext">The inner Microsoft.Bot.Builder.Dialogs.DialogContext for the current turn of conversation.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        protected override async Task<DialogTurnResult> OnContinueDialogAsync(DialogContext dialogContext, CancellationToken cancellationToken = default)
        {
            var result = await this.InterruptAsync(dialogContext, cancellationToken).ConfigureAwait(false);
            if (result != null)
            {
                return result;
            }

            return await base.OnContinueDialogAsync(dialogContext, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Handling interruption.
        /// </summary>
        /// <param name="dialogContext">The inner Microsoft.Bot.Builder.Dialogs.DialogContext for the current turn of conversation.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private async Task<DialogTurnResult> InterruptAsync(DialogContext dialogContext, CancellationToken cancellationToken = default)
        {
            if (dialogContext.Context.Activity.Type != ActivityTypes.Message)
            {
                return null;
            }

            var text = dialogContext.Context.Activity.Text;
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            if (LogoutCommands.Contains(text.ToUpperInvariant().Trim()))
            {
                // The bot adapter encapsulates the authentication processes.
                var botAdapter = (BotFrameworkAdapter)dialogContext.Context.Adapter;
                await botAdapter.SignOutUserAsync(dialogContext.Context, this.connectionName, null, cancellationToken).ConfigureAwait(false);
                await dialogContext.Context.SendActivityAsync(MessageFactory.Text(Strings.SignOutText), cancellationToken).ConfigureAwait(false);
                return await dialogContext.CancelAllDialogsAsync(cancellationToken).ConfigureAwait(false);
            }

            return null;
        }
    }
}