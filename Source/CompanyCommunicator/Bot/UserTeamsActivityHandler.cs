// <copyright file="UserTeamsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RepliesData;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Company Communicator User Bot.
    /// Captures user data, team data.
    /// </summary>
    public class UserTeamsActivityHandler : TeamsActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";
        private static readonly string AdaptiveCardContentType = "application/vnd.microsoft.card.adaptive";

        private readonly TeamsDataCapture teamsDataCapture;
        private readonly ISendingNotificationDataRepository notificationDataRepository;
        private readonly IReplyDataRepository replyDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserTeamsActivityHandler"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        /// <param name="notificationDataRepository">notificationDataRepository.</param>
        /// <param name="replyDataRepository">replyDataRepository.</param>
        public UserTeamsActivityHandler(TeamsDataCapture teamsDataCapture, ISendingNotificationDataRepository notificationDataRepository, IReplyDataRepository replyDataRepository)
        {
            this.teamsDataCapture = teamsDataCapture ?? throw new ArgumentNullException(nameof(teamsDataCapture));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.replyDataRepository = replyDataRepository ?? throw new ArgumentNullException(nameof(replyDataRepository));
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            if (activity.MembersAdded != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(activity);
            }

            if (activity.MembersRemoved != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var notificationId = (turnContext.Activity.Value as JObject)?["notificationId"]?.ToString();
            var action = (turnContext.Activity.Value as JObject)?["action"]?.ToString();
            var chatMessageId = turnContext.Activity.ReplyToId;

            switch (action)
            {
                case "getReplyCard":
                    var notification = await this.notificationDataRepository.GetAsync(NotificationDataTableNames.SendingNotificationsPartition, notificationId);
                    var content = JsonConvert.DeserializeObject<JObject>(notification.Content);
                    var adaptiveCardAttachment = new Attachment()
                    {
                        ContentType = AdaptiveCardContentType,
                        Content = content,
                    };
                    var activity = MessageFactory.Attachment(adaptiveCardAttachment);
                    activity.Id = chatMessageId;
                    await turnContext.UpdateActivityAsync(activity, cancellationToken);

                    var replyInputCard = this.GetReplyInputCard(notificationId);
                    await turnContext.SendActivityAsync(replyInputCard, cancellationToken);
                    break;
                case "saveReply":
                    await turnContext.DeleteActivityAsync(chatMessageId, cancellationToken);
                    var comment = (turnContext.Activity.Value as JObject)?["comment"]?.ToString();
                    var replyEntity = new ReplyDataEntity { PartitionKey = notificationId, RowKey = turnContext.Activity.From.AadObjectId, NotificationId = notificationId, AuthorId = turnContext.Activity.From.AadObjectId, AuthorDisplayName = turnContext.Activity.From.Name, Comment = comment };
                    await this.replyDataRepository.InsertOrMergeAsync(replyEntity);
                    break;
                default:
                    break;
            }
        }

        private IActivity GetReplyInputCard(string notificationId)
        {
            var content = JsonConvert.DeserializeObject<JObject>($"{{ \"type\": \"AdaptiveCard\", \"$schema\": \"{AdaptiveCardContentType}\", \"version\": \"1.2\", \"body\": [ {{ \"type\": \"Input.Text\", \"id\": \"comment\", \"placeholder\": \"Enter your comment\", \"isMultiline\": true }} ], \"actions\": [ {{ \"type\": \"Action.Submit\", \"title\": \"Send\", \"data\": {{ \"msteams\": {{ \"type\": \"messageBack\", \"value\": {{ \"notificationId\" : \"{notificationId}\", \"action\": \"saveReply\",  }} }} }}  }} ] }}");
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                Content = content,
            };
            return MessageFactory.Attachment(adaptiveCardAttachment);
        }

        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return UserTeamsActivityHandler.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}