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
        /// Instance of the Microsoft Bot Connector OAuthClient class.
        /// </summary>
        private readonly OAuthClient oAuthClient;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly BotSettings botSettings;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenHelper"/> class.
        /// Helps generating custom token, validating custom token and generate AADv1 user access token for given resource.
        /// </summary>
        /// <param name="oAuthClient">Instance of the Microsoft Bot Connector OAuthClient class.</param>
        /// <param name="botSettings">A set of key/value application configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TokenHelper(OAuthClient oAuthClient, IOptionsMonitor<BotSettings> botSettings, ILogger<TokenHelper> logger)
        {
            this.botSettings = botSettings.CurrentValue;
            this.oAuthClient = oAuthClient;
            this.logger = logger;
        }

        /// <inheritdoc/>
        public string GenerateAPIAuthToken(string aadObjectId, string serviceURL, string fromId, int jwtExpiryMinutes)
        {
            SymmetricSecurityKey signingKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.botSettings.TokenSigningKey));
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
                Issuer = this.botSettings.AppBaseUri,
                Audience = this.botSettings.AppBaseUri,
                IssuedAt = DateTime.UtcNow,
                Expires = DateTime.UtcNow.AddMinutes(jwtExpiryMinutes),
            };

            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            SecurityToken token = tokenHandler.CreateToken(securityTokenDescriptor);

            return tokenHandler.WriteToken(token);
        }

        /// <inheritdoc/>
        public async Task<string> GetUserTokenAsync(string fromId, string resourceUrl)
        {
            try
            {
                var token = await this.oAuthClient.UserToken.GetAadTokensAsync(fromId, this.botSettings.OAuthConnectionName, new Bot.Schema.AadResourceUrls { ResourceUrls = new string[] { resourceUrl } }).ConfigureAwait(false);
                return token?[resourceUrl]?.Token;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Failed to get user AAD access token for given resource using bot OAuthClient instance.");
                return null;
            }
        }
    }
}
