

// See https://aka.ms/new-console-template for more information
using GoogleAdapter.Adapters;

Console.WriteLine("Reading and writing to a Google spreadsheet...");
string secretsJsonPath = args[0];
string documentId = args[1];
string range = args[2];
string targetDocumentId = args[3];
string targetRange = args[4];

string json = File.ReadAllText(secretsJsonPath);

var sheets = new Sheets(json, isServiceAccount: true);

IList<IList<object>> values = sheets.Read(documentId: documentId, range: range);

//var values = new[] { new[] { (object)DateTime.Now.ToLongTimeString(), (object)"Azure Function" } };

sheets.Write(
    documentId: targetDocumentId,
    range: targetRange,
    values: values);

//IList<IList<object>> linesToEdit = tester.Read(documentId: documentId, range: range);

//for (int r = 4; r < 7; r++)
//{
//    linesToEdit[r][2] = "2/3/1111";
//}

//var oValues = new[] { new[] { (object)"4/4/4444" } };

//tester.Write(
//    documentId: documentId,
//    range: range,
//    values: oValues); 

//IList<IList<object>> lines = tester.Read(documentId: documentId, range: range);

//foreach(IList<object> line in lines)
//{
//    foreach(object value in line)
//    {
//        Console.Write($"{value}, ");
//    }
//    Console.WriteLine();
//}
