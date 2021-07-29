// <copyright file="IDriveItemsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using Microsoft.AspNetCore.Http;
    using Microsoft.Graph;
    using System.IO;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Drive Items Service.
    /// </summary>
    public interface IDriveItemsService
    {
        /// <summary>
        /// upload file for group.
        /// </summary>
        /// <param name="groupId">siteId.</param>
        /// <param name="file">file.</param>
        /// <param name="fileName">fileName.</param>
        /// <returns>file url.</returns>
        Task<string> UploadFileForGroup(string groupId, IFormFile file, string fileName);

        /// <summary>
        /// get file stream by rel path.
        /// </summary>
        /// <param name="path">path.</param>
        /// <param name="groupId">groupId.</param>
        /// <returns>file.</returns>
        Task<Stream> GetFileStreamByPath(string path, string groupId);

        /// <summary>
        /// get file stream by rel path.
        /// </summary>
        /// <param name="path">path.</param>
        /// <param name="groupId">groupId.</param>
        /// <returns>file.</returns>
        Task<DriveItem> GetFileByPath(string path, string groupId);
    }
}
