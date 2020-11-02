using Google.Sheets.Apis.Utility;
using Google.Sheets.Apis.Utility.Sheets.v4;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Server.Settings;

namespace Server.AppStart
{
    public static class ServiceCollectionExtension
    {
        public static void AddAppSettingSetup(this IServiceCollection service, IConfiguration configuration)
        {
            service.AddSingleton(c => configuration.GetConfig<GoogleSheetsApi>());
        }

        public static void AddGoogleSheetsV4ApiSetup(this IServiceCollection service)
        {
            // Create GoogleCredentialFactory
            service.AddSingleton<GoogleCredentialFactory>();

            // Create GoogleCredentialFactory
            service.AddSingleton(c => new SheetsServiceFactory(c.GetService<GoogleCredentialFactory>()));

            // Create SheetsService
            service.AddScoped(c => c.GetService<SheetsServiceFactory>().Create(c.GetService<GoogleSheetsApi>().ServiceAccountName));
        }

        public static void AddSwaggerSetup(this IServiceCollection services)
        {
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo { Title = "Google Sheet Api", Version = "v1" });
            });
        }

        public static void AddLoggingSetup(this IServiceCollection services)
        {
            // 設定Log輸出到Console畫面上，可以設定log的顯示等級
            // 此方式會輸出到docker logs上
            services.AddLogging(builder =>
            {
                builder.AddDebug()
                  .AddConsole()
                  //.AddConfiguration(_configuration.GetSection("Logging"))
                  .SetMinimumLevel(LogLevel.Information);
            });
        }
    }
}
