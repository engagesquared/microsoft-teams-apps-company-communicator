// <copyright file="ReplyDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReplyData
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RepliesData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;

    /// <summary>
    /// Repository of the team data stored in the table storage.
    /// </summary>
    public class ReplyDataRepository : BaseRepository<ReplyDataEntity>, IReplyDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReplyDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public ReplyDataRepository(
            ILogger<ReplyDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: ReplyDataTableNames.TableName,
                  defaultPartitionKey: ReplyDataTableNames.TeamDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<ReplyDataEntity>> GetAllRepliesByNotificationId(string notificationId)
        {
            if (notificationId == null)
            {
                return null;
            }

            var notificationFilter = TableQuery.GenerateFilterCondition("NotificationId", QueryComparisons.Equal, notificationId);
            var replies = await this.GetWithFilterAsync(notificationFilter, notificationId);
            return replies;
        }
    }
}
