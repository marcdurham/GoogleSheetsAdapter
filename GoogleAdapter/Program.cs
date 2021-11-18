

// See https://aka.ms/new-console-template for more information
using GoogleAdapter;

Console.WriteLine("Reading and writing to a Google spreadsheet...");
string secretsJsonPath = args[0];
string documentId = args[1];
string range = args[2];

var tester = new Sheets(secretsJsonPath);

IList<IList<object>> linesToEdit = tester.Read(documentId: documentId, range: range);

//for (int r = 4; r < 7; r++)
//{
//    linesToEdit[r][2] = "2/3/1111";
//}

var oValues = new[] { new[] { (object)"2/4/1111" } };

tester.Write(
    documentId: documentId,
    range: range,
    values: oValues); // new[] { new[] { (object)"2/5/1111" } });

IList<IList<object>> lines = tester.Read(documentId: documentId, range: range);

foreach(IList<object> line in lines)
{
    foreach(object value in line)
    {
        Console.Write($"{value}, ");
    }
    Console.WriteLine();
}
