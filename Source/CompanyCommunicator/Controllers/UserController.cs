// <copyright file="UserController.cs" company="Engage Squared">
// Copyright (c) Engage Squared. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Identity.Web;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;

    [Route("api/user")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UserController : ControllerBase
    {
        private readonly ITokenAcquisition tokenAcquisition;

        public UserController(ITokenAcquisition tokenAcquisition)
        {
            this.tokenAcquisition = tokenAcquisition ?? throw new ArgumentNullException(nameof(tokenAcquisition));
        }

        [HttpGet("GetToken")]
        public async Task<IActionResult> GetToken()
        {
            try
            {
                var accessToken = await this.tokenAcquisition.GetAccessTokenForUserAsync(new string[] { Common.Constants.ScopeDefault });
                return this.Ok(accessToken);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
