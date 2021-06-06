// <copyright file="DriveItemsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    /// <summary>
    /// Groups Service.
    /// </summary>
    internal class DriveItemsService : IDriveItemsService
    {
        private readonly IGraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="DriveItemsService"/> class.
        /// </summary>
        /// <param name="graphServiceClient">graph service client.</param>
        internal DriveItemsService(IGraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        private int MaxResultCount { get; set; } = 25;

        private int MaxRetry { get; set; } = 2;

        /// <summary>
        /// get groups by ids.
        /// </summary>
        /// <param name="siteId">siteId.</param>
        /// <param name="webId">webId.</param>
        /// <param name="listId">listId.</param>
        /// <param name="sharepointHostName">sharepointHostName.</param>
        /// <param name="stream">stream.</param>
        /// <returns>file url.</returns>
        public async Task<string> UploadFileToPublicCDN(string siteId, string webId, string listId, string sharepointHostName, Stream stream, string fileName)
        {
            var result = await this.graphServiceClient
                            .Sites[$"{sharepointHostName}.sharepoint.com,{siteId},{webId}"]
                            .Lists[listId]
                            .Drive
                            .Root
                            .ItemWithPath(fileName)
                            .Content
                            .Request()
                            .PutAsync<DriveItem>(stream);
            return result.WebUrl;
        }
    }
}
