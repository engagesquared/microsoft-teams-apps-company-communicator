// <copyright file="IRepliesDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RepliesData
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Team Data Repository.
    /// </summary>
    public interface IReplyDataRepository : IRepository<ReplyDataEntity>
    {
        /// <summary>
        /// Gets replies.
        /// </summary>
        /// <param name="notificationId">notification id.</param>
        /// <returns>Team data entities.</returns>
        public Task<IEnumerable<ReplyDataEntity>> GetAllRepliesByNotificationId(string notificationId);
    }
}
