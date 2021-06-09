// <copyright file="ReplyDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RepliesData
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// ReplyDataEntity entity class.
    /// </summary>
    public class ReplyDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets Comment.
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// Gets or sets NotificationId.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets AuthorId.
        /// </summary>
        public string AuthorId { get; set; }

        /// <summary>
        /// Gets or sets AuthorDisplayName.
        /// </summary>
        public string AuthorDisplayName { get; set; }
    }
}
