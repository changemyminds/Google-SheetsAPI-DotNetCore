using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Server.Settings
{
    public static class AppSettingsExtension
    {
        public static TConfig GetConfig<TConfig>(this IConfiguration configuration, string section = "")
            where TConfig : class, new()
        {
            var config = new TConfig();
            configuration.GetSection(string.IsNullOrEmpty(section) ? typeof(TConfig).Name : section)
                .Bind(config);
            return config;
        }

    }
}
