// <copyright file="HistoryNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Identity.Web;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the notifications history data.
    /// </summary>
    [Route("api/history")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class HistoryNotificationsController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IGroupsService groupsService;

        /// <summary>
        /// Initializes a new instance of the <see cref="HistoryNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="groupsService">group service.</param>
        public HistoryNotificationsController(
            INotificationDataRepository notificationDataRepository,
            IGroupsService groupsService)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.groupsService = groupsService ?? throw new ArgumentNullException(nameof(groupsService));
        }

        /// <summary>
        /// Get user notifications history by user Id.
        /// </summary>
        /// <returns>It returns the notification history for specified user.</returns>
        [HttpGet]
        public async Task<ActionResult<DraftNotification>> GetUserNotificationsHistoryByIdAsync()
        {
            var claim = this.User?.Claims?.FirstOrDefault(p => p.Type == ClaimConstants.ObjectId);
            var userAadId = claim?.Value;

            var userGroupsRes = await this.groupsService.GetUserGroups(userAadId);
            var groupIds = userGroupsRes.ToList().Select(x => x.Id);
            var result = await this.notificationDataRepository.GetSentNotificationsToUser(groupIds);
            return this.Ok(result);
        }
    }
}
