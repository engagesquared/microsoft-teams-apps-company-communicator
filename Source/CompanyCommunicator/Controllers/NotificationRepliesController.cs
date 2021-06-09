// <copyright file="NotificationRepliesController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RepliesData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Controllers.Options;

    [Route("api/notificationReplies")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class NotificationRepliesController : ControllerBase
    {
        private readonly bool isEnableReplyFunctionality;
        private readonly IReplyDataRepository replyDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationRepliesController"/> class.
        /// </summary>
        /// <param name="notificationReplyOptions">.</param>
        public NotificationRepliesController(IOptions<NotificationReplyOptions> notificationReplyOptions, IReplyDataRepository replyDataRepository)
        {
            if (notificationReplyOptions is null)
            {
                throw new ArgumentNullException(nameof(notificationReplyOptions));
            }

            this.isEnableReplyFunctionality = notificationReplyOptions.Value.EnableReplyFunctionality;
            this.replyDataRepository = replyDataRepository;
        }

        /// <summary>
        /// Retrieve EnableReplyFunctionality value.
        /// </summary>
        /// <returns>data</returns>
        [HttpGet("isEnableReplyFunctionality")]
        public ActionResult<bool> GetPublicCDNLibraryDataForMGTControl()
        {
            return this.Ok(this.isEnableReplyFunctionality);
        }

        /// <summary>
        /// get replies in csv format.
        /// </summary>
        /// <returns>file url</returns>
        [HttpGet("replies/{id}")]
        public async Task<ActionResult> UploadFile(string id)
        {
            var replies = await this.replyDataRepository.GetAllRepliesByNotificationId(id);
            var content = this.GetCsvReport(replies.ToList());
            string file_type = "text/csv";
            string file_name = $"Replies-{id}-{DateTime.Now.ToString("yyyyMMdd_HHmm")}.csv";
            var file = File(content, file_type, file_name);
            return this.Ok(file);
        }

        public byte[] GetCsvReport(List<ReplyDataEntity> replies)
        {
            var report = new StringBuilder();
            report.AppendLine("Replies:");
            report.AppendLine("AuthorId;AuthorDisplayName;Comment");
            replies.ForEach(x => report.AppendLine($"{x.AuthorId};{x.AuthorDisplayName};{x.Comment}"));
            var file = Encoding.UTF8.GetBytes(report.ToString());
            return file;
        }
    }
}
