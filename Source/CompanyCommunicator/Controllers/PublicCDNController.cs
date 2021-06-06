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
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Controllers.Options;

    [Route("api/publicCDN")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class PublicCDNController : ControllerBase
    {
        private readonly IDriveItemsService driveItemsService;
        private readonly string sharepointHostName;
        private readonly string siteId;
        private readonly string webId;
        private readonly string libraryId;

        /// <summary>
        /// Initializes a new instance of the <see cref="PublicCDNController"/> class.
        /// </summary>
        /// <param name="publicCDNOptions">The authentication options.</param>
        public PublicCDNController(IOptions<PublicCDNOptions> publicCDNOptions, IDriveItemsService driveItemsService)
        {
            if (publicCDNOptions is null)
            {
                throw new ArgumentNullException(nameof(publicCDNOptions));
            }

            this.driveItemsService = driveItemsService;
            this.sharepointHostName = publicCDNOptions.Value.SharepointHostName;
            this.siteId = publicCDNOptions.Value.SiteId;
            this.webId = publicCDNOptions.Value.WebId;
            this.libraryId = publicCDNOptions.Value.LibraryId;
        }

        /// <summary>
        /// Retrieve public cdn options.
        /// </summary>
        /// <returns>data</returns>
        [HttpGet("options")]
        public ActionResult<PublicCDNOptions> GetPublicCDNLibraryDataForMGTControl()
        {
            var result = new PublicCDNOptions() { LibraryId = this.libraryId, WebId = this.webId, SiteId = this.siteId, SharepointHostName = this.sharepointHostName } as object;
            return this.Ok(result);
        }

        /// <summary>
        /// Upload file.
        /// </summary>
        /// <returns>file url</returns>
        [HttpPost("content")]
        public async Task<ActionResult<string>> UploadFile(IFormFile file)
        {
            var stream = file.OpenReadStream();
            var link = await this.driveItemsService.UploadFileToPublicCDN(this.siteId, this.webId, this.libraryId, this.sharepointHostName, stream, file.FileName);
            return this.Ok(link);
        }
    }
}
