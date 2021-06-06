// <copyright file="IDriveItemsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System.IO;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Drive Items Service.
    /// </summary>
    public interface IDriveItemsService
    {
        /// <summary>
        /// get groups by ids.
        /// </summary>
        /// <param name="siteId">siteId.</param>
        /// <param name="webId">webId.</param>
        /// <param name="listId">listId.</param>
        /// <param name="sharepointHostName">sharepointHostName.</param>
        /// <param name="stream">stream.</param>
        /// <param name="fileName">fileName.</param>
        /// <returns>file url.</returns>
        Task<string> UploadFileToPublicCDN(string siteId, string webId, string listId, string sharepointHostName, Stream stream, string fileName);
    }
}
