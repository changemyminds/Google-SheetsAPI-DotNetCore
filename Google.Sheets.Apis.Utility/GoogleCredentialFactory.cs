using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Google.Sheets.Apis.Utility
{
    public class GoogleCredentialFactory
    {
        public GoogleCredential Create(string serviceAccountPath)
        {
            using (var stream = new FileStream(serviceAccountPath, FileMode.Open, FileAccess.Read))
            {
                return GoogleCredential.FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
        }
    }
}
