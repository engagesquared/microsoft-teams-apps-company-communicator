// <copyright file="DriveItemsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
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
        /// get file by rel path.
        /// </summary>
        /// <param name="path">path.</param>
        /// <param name="groupId">groupId.</param>
        /// <returns>file.</returns>
        public async Task<Stream> GetFileStreamByPath(string path, string groupId)
        {
            var result = await this.graphServiceClient.Groups[groupId]
                                                    .Drive
                                                    .Root
                                                    .ItemWithPath(path)
                                                    .Content
                                                    .Request()
                                                    .GetAsync();
            return result;
        }

        /// <summary>
        /// get file by rel path.
        /// </summary>
        /// <param name="path">path.</param>
        /// <param name="groupId">groupId.</param>
        /// <returns>file.</returns>
        public async Task<DriveItem> GetFileByPath(string path, string groupId)
        {
            var result = await this.graphServiceClient.Groups[groupId]
                                                    .Drive
                                                    .Root
                                                    .ItemWithPath(path)
                                                    .Request()
                                                    .GetAsync();
            return result;
        }

        /// <summary>
        /// upload file for group.
        /// </summary>
        /// <param name="groupId">groupId.</param>
        /// <param name="stream">stream.</param>
        /// <returns>file url.</returns>
        public async Task<string> UploadFileForGroup(string groupId, IFormFile file, string fileName)
        {
            var stream = file.OpenReadStream();
            var result = await this.graphServiceClient.Groups[groupId]
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
