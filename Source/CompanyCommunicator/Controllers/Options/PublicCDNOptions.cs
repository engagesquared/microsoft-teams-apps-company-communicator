// <copyright file="PublicCDNOptions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers.Options
{
    /// <summary>
    /// Public CDN app options.
    /// </summary>
    public class PublicCDNOptions
    {
        /// <summary>
        /// Gets or sets a value of SharepointHostName.
        /// </summary>
        public string SharepointHostName { get; set; }

        /// <summary>
        /// Gets or sets a value of SiteId.
        /// </summary>
        public string SiteId { get; set; }

        /// <summary>
        /// Gets or sets a value of LibraryId.
        /// </summary>
        public string LibraryId { get; set; }

        /// <summary>
        /// Gets or sets a value of WebId.
        /// </summary>
        public string WebId { get; set; }
    }
}
