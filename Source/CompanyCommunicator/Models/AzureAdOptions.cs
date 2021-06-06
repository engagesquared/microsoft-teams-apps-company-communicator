﻿// <copyright file="AzureAdOptions.cs" company="Engage Squared">
// Copyright (c) Engage Squared. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    /// <summary>
    /// AzureAdOptions class contain value application configuration properties for Azure Active Directory.
    /// </summary>
    public class AzureAdOptions
    {
        /// <summary>
        /// Gets or sets Client Id.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets Client secret.
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Gets or sets Graph API scope.
        /// </summary>
        public string GraphScope { get; set; }

        /// <summary>
        /// Gets or sets Application Id URI.
        /// </summary>
        public string ApplicationIdUri { get; set; }

        /// <summary>
        /// Gets or sets valid isuers.
        /// </summary>
        public string ValidIssuers { get; set; }

        /// <summary>
        /// Gets or sets tenant Id.
        /// </summary>
        public string TenantId { get; set; }
    }
}
