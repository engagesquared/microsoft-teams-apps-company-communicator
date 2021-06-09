// <copyright file="ReplyDataTableNames.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData
{
    /// <summary>
    /// Replies data table names.
    /// </summary>
    public static class ReplyDataTableNames
    {
        /// <summary>
        /// Table name for the Replies data table.
        /// </summary>
        public static readonly string TableName = "Replies";

        /// <summary>
        /// Replies data partition key name.
        /// </summary>
        public static readonly string TeamDataPartition = "Replies";
    }
}
