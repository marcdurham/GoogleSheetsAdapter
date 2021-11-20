using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Text;

namespace GoogleAdapter.Adapters
{
    public class Sheets
    {
        // Some APIs, like Storage, accept a credential in their Create()
        // method.
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly, SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Experiment";
        const string CredPath = "token.json";

        readonly SheetsService _service;

        public Sheets(string json, bool isServiceAccount = false)
        {
            if (isServiceAccount)
            {
                _service = GetSheetServiceService(json);
            }
            else
            {
                _service = GetSheetService(json, U);
            }
        }

        public void WriteOneCell(string documentId, string range, object value)
        {
            Write(documentId, range, new[] { new[] { value } });
        }

        public void Write(string documentId, string range, IList<IList<object>> values)
        {
            ValueRange valueRange = new() { Values = values };
            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest =
                _service.Spreadsheets.Values.Update(valueRange, documentId, range);

            updateRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            updateRequest.Execute();
        }

        public IList<IList<object>> Read(string documentId, string range)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request =
                _service.Spreadsheets.Values.Get(documentId, range);

            ValueRange response = request.Execute();

            return response.Values;
        }

        static SheetsService GetSheetService(string jsonPath, ICredential credential)
        {
            //UserCredential credential = Credentials(jsonPath);

            // Create Google Sheets API service.
            SheetsService service = new(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        static SheetsService GetSheetServiceService(string json)
        {
            GoogleCredential credential = ServiceCredentials(json);

            // Create Google Sheets API service.
            SheetsService service = new(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }
        static UserCredential Credentials(string json)
        {
            UserCredential credential;
           
            byte[] byteArray = Encoding.ASCII.GetBytes(json);
            using (MemoryStream stream = new MemoryStream(byteArray))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.

                string path = Path.Combine(Path.GetTempPath(), CredPath);
                
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(path, true)).Result;
                
                Console.WriteLine("Credential file saved to: " + CredPath);
            }

            return credential;
        }

        static GoogleCredential ServiceCredentials(string json)
        {
            GoogleCredential credential;

            byte[] byteArray = Encoding.ASCII.GetBytes(json);
            using (MemoryStream stream = new MemoryStream(byteArray))
            {
                var serviceCred = ServiceAccountCredential.FromServiceAccountData(stream);
                credential = GoogleCredential.FromServiceAccountCredential(serviceCred);
            }

            return credential;
        }
    }
}
