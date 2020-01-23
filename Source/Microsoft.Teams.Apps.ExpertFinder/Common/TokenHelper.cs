// <copyright file="TokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder.Common
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Connector;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;

    /// <summary>
    /// Helper class for JWT token generation, validation and generate AAD user access token for given resource, e.g. Microsoft Graph, SharePoint.
    /// </summary>
    public class TokenHelper : ITokenHelper, ICustomTokenHelper
    {
        /// <summary>
        /// Random key to create jwt security key.
        /// </summary>
        private readonly string securityKey;

        /// <summary>
        /// Instance of the Microsoft Bot Connector OAuthClient class.
        /// </summary>
        private readonly OAuthClient oAuthClient;

        /// <summary>
        /// Application base uri.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// AADv1 bot connection name.
        /// </summary>
        private readonly string connectionName;

        /// <summary>
        /// Represents a set of key/value application configuration properties related to custom token.
        /// </summary>
        private readonly TokenSettings options;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenHelper"/> class.
        /// Helps generating custom token, validating custom token and generate AADv1 user access token for given resource.
        /// </summary>
        /// <param name="oAuthClient">Instance of the Microsoft Bot Connector OAuthClient class.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties jwt access token.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TokenHelper(OAuthClient oAuthClient, IOptionsMonitor<TokenSettings> optionsAccessor, ILogger<TokenHelper> logger)
        {
            this.options = optionsAccessor.CurrentValue;
            this.oAuthClient = oAuthClient;
            this.appBaseUri = this.options.AppBaseUri;
            this.securityKey = this.options.SecurityKey;
            this.connectionName = this.options.ConnectionName;
            this.logger = logger;
        }

        /// <summary>
        /// Generate custom jwt access token to authenticate/verify valid request on api side.
        /// </summary>
        /// <param name="aadObjectId">User account's object id within Azure Active Directory.</param>
        /// <param name="serviceURL">Service uri where responses to this activity should be sent.</param>
        /// <param name="fromId">Unique user id from activity.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token.</param>
        /// <returns>Custom jwt access token.</returns>
        public string GenerateAPIAuthToken(string aadObjectId, string serviceURL, string fromId, int jwtExpiryMinutes)
        {
            SymmetricSecurityKey signingKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.securityKey));
            SigningCredentials signingCredentials = new SigningCredentials(signingKey, SecurityAlgorithms.HmacSha256);

            SecurityTokenDescriptor securityTokenDescriptor = new SecurityTokenDescriptor()
            {
                Subject = new ClaimsIdentity(
                    new List<Claim>()
                    {
                        new Claim("aadObjectId", aadObjectId),
                        new Claim("serviceURL", serviceURL),
                        new Claim("fromId", fromId),
                    }, "Custom"),
                NotBefore = DateTime.UtcNow,
                SigningCredentials = signingCredentials,
                Issuer = this.appBaseUri,
                Audience = this.appBaseUri,
                IssuedAt = DateTime.UtcNow,
                Expires = DateTime.UtcNow.AddMinutes(jwtExpiryMinutes),
            };

            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            SecurityToken token = tokenHandler.CreateToken(securityTokenDescriptor);

            return tokenHandler.WriteToken(token);
        }

        /// <summary>
        /// Get user access token for given resource using Bot OAuth client instance.
        /// </summary>
        /// <param name="fromId">Activity from id.</param>
        /// <param name="resourceUrl">Resource url for which token will be acquired.</param>
        /// <returns>A task that represents security access token for given resource.</returns>
        public async Task<string> GetUserTokenAsync(string fromId, string resourceUrl)
        {
            try
            {
                var token = await this.oAuthClient.UserToken.GetAadTokensAsync(fromId, this.connectionName, new Bot.Schema.AadResourceUrls { ResourceUrls = new string[] { resourceUrl } }).ConfigureAwait(false);
                return token?[resourceUrl]?.Token;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Failed to get user AAD access token for given resource using Bot OAuth client instance.");
                return default;
            }
        }
    }
}
