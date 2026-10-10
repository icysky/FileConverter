#r "nuget: DocumentFormat.OpenXml, 3.3.0"

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var path = Args[0];
var doc = WordprocessingDocument.Open(path, false);
var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
int count = 0;
foreach (var error in validator.Validate(doc))
{
    count++;
    Console.WriteLine(string.Format("{0}: {1}", error.ErrorType, error.Description));
    Console.WriteLine(string.Format("  at {0}", error.Path == null ? "" : error.Path.XPath));
}
Console.WriteLine(string.Format("Validation errors: {0}", count));
doc.Dispose();
