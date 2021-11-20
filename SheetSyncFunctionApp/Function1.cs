using GoogleAdapter.Adapters;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace SheetSyncFunctionApp
{
    public class Function1
    {
        [FunctionName("Function1")]
        public void Run([TimerTrigger("47 */5 * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            // See https://aka.ms/new-console-template for more information

            log.LogInformation("Reading and writing to a Google spreadsheet...");
            string secretsJson = Environment.GetEnvironmentVariable("ServiceSecretsJson", EnvironmentVariableTarget.Process);
            string documentId = Environment.GetEnvironmentVariable("DocumentId", EnvironmentVariableTarget.Process);
            string range = Environment.GetEnvironmentVariable("Range", EnvironmentVariableTarget.Process);
            string targetDocumentId = Environment.GetEnvironmentVariable("TargetDocumentId", EnvironmentVariableTarget.Process);
            string targetRange = Environment.GetEnvironmentVariable("TargetRange", EnvironmentVariableTarget.Process);
            string sendGridApiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY", EnvironmentVariableTarget.Process);

            CopySheet(log, secretsJson, documentId, range, targetDocumentId, targetRange);

            var sheets = new Sheets(secretsJson, isServiceAccount: true);

            IList<IList<object>> rows = sheets.Read(documentId: documentId, range: range);

            foreach (IList<object> row in rows)
            {
                string name = $"{row[0]}";
                string email = $"{row[1]}";
                log.LogInformation($"Sending email to {name}: {email} ...");
                Response response = SendEmail(name, email, sendGridApiKey).Result;
                log.LogInformation($"Status Code:{response.StatusCode}");
            }
        }

        private static async Task<Response> SendEmail(string name, string email, string sendGridApiKey)
        {
            //HttpClient client = new();
            //HttpRequestMessage request = new(HttpMethod.Get, "https://api.sendgrid.com/v3/mail/send");
            //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", sendGridApiKey);
            //client.DefaultRequestHeaders.
            //client.SendAsync(request);
            ////var apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
            var client = new SendGridClient(sendGridApiKey);
            var from = new EmailAddress("auto@territorytools.org", "Territory Tools System");
            var subject = "Test Email";
            var to = new EmailAddress(email, name);
            var plainTextContent = "and easy to do anywhere, even with C#";
            var htmlContent = "<strong>and easy to do anywhere, even with C#</strong>";
            var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);
            var response = await client.SendEmailAsync(msg).ConfigureAwait(false);
            return response;
        }

        private static void CopySheet(ILogger log, string secretsJson, string documentId, string range, string targetDocumentId, string targetRange)
        {
            log.LogInformation($"secretsJsonPath:{secretsJson.Length} documentId:{documentId} range:{range}");

            var sheets = new Sheets(secretsJson, isServiceAccount: true);

            IList<IList<object>> values = sheets.Read(documentId: documentId, range: range);

            //var values = new[] { new[] { (object)DateTime.Now.ToLongTimeString(), (object)"Azure Function" } };

            sheets.Write(
                documentId: targetDocumentId,
                range: targetRange,
                values: values);
        }
    }
}
