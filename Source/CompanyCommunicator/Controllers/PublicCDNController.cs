// <copyright file="PublicCDNController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AzureStorage;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;

    [Route("api/publicCDN")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class PublicCDNController : ControllerBase
    {
        private readonly IDriveItemsService driveItemsService;
        private readonly CDNImagesBlobContainerService cdnImagesBlobContainerClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="PublicCDNController"/> class.
        /// </summary>
        public PublicCDNController(IDriveItemsService driveItemsService, CDNImagesBlobContainerService cdnImagesBlobContainerClient)
        {
            this.driveItemsService = driveItemsService;
            this.cdnImagesBlobContainerClient = cdnImagesBlobContainerClient;
        }


        /// <summary>
        /// Create public copy.
        /// </summary>
        /// <returns>file url</returns>
        [HttpPost("copy")]
        public async Task<ActionResult<string>> UploadFile(string path, string groupId)
        {
            var driveItemTask = this.driveItemsService.GetFileByPath(path, groupId);
            var driveItemStreamTask = this.driveItemsService.GetFileStreamByPath(path, groupId);
            await Task.WhenAll(driveItemTask, driveItemStreamTask);
            var driveItem = driveItemTask.Result;
            var driveItemStream = driveItemStreamTask.Result;
            var file = new FormFile(driveItemStream, 0, driveItemStream.Length, driveItem.Name, driveItem.Name);
            var fileAbsoluteUrl = await this.cdnImagesBlobContainerClient.UploadFileToBlobContainer(file, $"{groupId}-{file.FileName}", driveItem.File.MimeType);
            return this.Ok(fileAbsoluteUrl);
        }

        /// <summary>
        /// Upload file.
        /// </summary>
        /// <returns>file url</returns>
        [HttpPost("content")]
        public async Task<ActionResult<string>> UploadFile(IFormFile file, string groupId)
        {
            await this.driveItemsService.UploadFileForGroup(groupId, file, file.FileName);
            var fileAbsoluteUrl = await this.cdnImagesBlobContainerClient.UploadFileToBlobContainer(file, $"{groupId}-{file.FileName}", file.ContentType);
            return this.Ok(fileAbsoluteUrl);
        }
    }
}
