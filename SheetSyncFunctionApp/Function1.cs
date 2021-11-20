using GoogleAdapter.Adapters;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;

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

            log.LogInformation($"secretsJsonPath:{secretsJson.Length} documentId:{documentId} range:{range}");

            var sheets = new Sheets(secretsJson, isServiceAccount: true);

            //IList<IList<object>> linesToEdit = tester.Read(documentId: documentId, range: range);

            var values = new[] { new[] { (object)DateTime.Now.ToLongTimeString(), (object)"Azure Function" } };

            sheets.Write(
                documentId: documentId,
                range: range,
                values: values);

            //IList<IList<object>> lines = tester.Read(documentId: documentId, range: range);

            //foreach (IList<object> line in lines)
            //{
            //    var builder = new StringBuilder(3000);
            //    foreach (object value in line)
            //    {
            //        builder.Append($"{value}, ");
            //    }

            //    log.LogInformation(builder.ToString());
            //}
        }
    }
}
