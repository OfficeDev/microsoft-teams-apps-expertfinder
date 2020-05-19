// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.ExpertFinder
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.SpaServices.ReactDevelopmentServer;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.ExpertFinder.Bots;
    using Microsoft.Teams.Apps.ExpertFinder.Common;
    using Microsoft.Teams.Apps.ExpertFinder.Common.Interfaces;
    using Microsoft.Teams.Apps.ExpertFinder.Models.Configuration;
    using Polly;
    using Polly.Extensions.Http;

    /// <summary>
    /// This a Startup class for this Bot.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">object that passes the application configuration key-values.</param>
        public Startup(IConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        /// <summary>
        /// Gets object that passes the application configuration key-values..
        /// </summary>
        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services"> Service Collection Interface.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddHttpClient<ISharePointApiHelper, SharePointApiHelper>().AddPolicyHandler(GetRetryPolicy());
            services.AddHttpClient<IGraphApiHelper, GraphApiHelper>().AddPolicyHandler(GetRetryPolicy());
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
            services.Configure<BotSettings>(options =>
            {
                options.AppBaseUri = this.Configuration["AppBaseUri"];
                options.AppInsightsInstrumentationKey = this.Configuration["APPINSIGHTS_INSTRUMENTATIONKEY"];
                options.OAuthConnectionName = this.GetFirstSetting("OAuthConnectionName", "ConnectionName");
                options.SharePointSiteUrl = this.Configuration["SharePointSiteUrl"];
                options.TenantId = this.Configuration["TenantId"];
                options.TokenSigningKey = this.GetFirstSetting("TokenSigningKey", "SecurityKey");
                options.StorageConnectionString = this.Configuration["StorageConnectionString"];
            });
            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddJwtBearer(options =>
                {
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateAudience = true,
                        ValidAudiences = new List<string> { this.Configuration["AppBaseUri"] },
                        ValidIssuers = new List<string> { this.Configuration["AppBaseUri"] },
                        ValidateIssuer = true,
                        ValidateIssuerSigningKey = true,
                        IssuerSigningKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.GetFirstSetting("TokenSigningKey", "SecurityKey"))),
                        RequireExpirationTime = true,
                        ValidateLifetime = true,
                        ClockSkew = TimeSpan.FromSeconds(30),
                    };
                });

            // In production, the React files will be served from this directory
            services.AddSpaStaticFiles(configuration =>
            {
                configuration.RootPath = "ClientApp/build";
            });

            services.AddMemoryCache();

            // Create the Bot Framework Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();

            services.AddSingleton(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"]));

            // For conversation state.
            services.AddSingleton<IStorage>(new AzureBlobStorage(this.Configuration["StorageConnectionString"], "bot-state"));

            // Create the Conversation state. (Used by the Dialog system itself.)
            services.AddSingleton<ConversationState>();

            // Create the User state. (Used in this bot's Dialog implementation.)
            services.AddSingleton<UserState>();

            // Create the telemetry middleware(used by the telemetry initializer) to track conversation events
            services.AddSingleton<TelemetryLoggerMiddleware>();

            // The Dialog that will be run by the bot.
            services.AddSingleton<MainDialog>();

            services.AddSingleton<ITokenHelper, TokenHelper>();
            services.AddSingleton<ICustomTokenHelper, TokenHelper>();
            services.AddSingleton<IUserProfileActivityStorageHelper, UserProfileActivityStorageHelper>();
            services.AddSingleton(new OAuthClient(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"])));

            // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
            services.AddTransient<IBot, ExpertFinderBot>();
            services.AddApplicationInsightsTelemetry();
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">Application Builder.</param>
        /// <param name="env">Hosting Environment.</param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseAuthentication();
            app.UseStaticFiles();
            app.UseSpaStaticFiles();
            app.UseMvc();
            app.UseStaticFiles();
            app.UseSpa(spa =>
            {
                spa.Options.SourcePath = "ClientApp";

                if (env.IsDevelopment())
                {
                    spa.UseReactDevelopmentServer(npmScript: "start");
                }
            });
        }

        /// <summary>
        /// Retry policy for for transient error cases.
        /// If there is no success code in response, request will be sent again for two times
        /// with interval of 2 and 8 seconds respectively.
        /// </summary>
        /// <returns>Policy.</returns>
        private static IAsyncPolicy<HttpResponseMessage> GetRetryPolicy()
        {
            return HttpPolicyExtensions
                .HandleTransientHttpError()
                .OrResult(response => response.IsSuccessStatusCode == false)
                .WaitAndRetryAsync(3, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)));
        }

        /// <summary>
        /// Find the first configuration value in the list that is not null or empty.
        /// </summary>
        /// <param name="keys">List of keys to check.</param>
        /// <returns>First configuration value that is not null or empty.</returns>
        private string GetFirstSetting(params string[] keys)
        {
            return keys.Select(key => this.Configuration[key]).Where(value => !string.IsNullOrEmpty(value)).FirstOrDefault();
        }
    }
}
