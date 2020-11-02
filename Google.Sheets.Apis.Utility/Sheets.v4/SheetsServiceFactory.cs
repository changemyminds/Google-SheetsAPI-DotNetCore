using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.IO;
using System.Reflection;

namespace Google.Sheets.Apis.Utility.Sheets.v4
{
    public class SheetsServiceFactory
    {
        private readonly GoogleCredentialFactory _factory;

        public SheetsServiceFactory(GoogleCredentialFactory factory)
        {
            _factory = factory;
        }

        public SheetsService Create(string serviceAccountName)
        {
            var directory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            
            var path = Path.Combine(directory, serviceAccountName);
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _factory.Create(path)
            });
        }
    }
}
