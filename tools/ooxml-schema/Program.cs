// Open XML SDK schema validation — the authoritative OOXML schema check that
// PowerPoint itself effectively uses. This is the optional "Layer 2" of the
// marp-native-pptx validator; the pure-Node validator (src/validate.ts) covers
// the PowerPoint-clean rules the schema does NOT express.
//
// Usage:  dotnet run --project tools/ooxml-schema -- <file.pptx>
// Prints one line per schema error and a final "TOTAL ERRORS: N".
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length < 1)
{
    Console.Error.WriteLine("usage: OoxmlSchema <file.pptx>");
    return 2;
}

using var doc = PresentationDocument.Open(args[0], false);
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
int n = 0;
foreach (var e in validator.Validate(doc))
{
    n++;
    Console.WriteLine($"[{n}] {e.Part?.Uri} :: {e.Path?.XPath} :: {e.Description}");
}
Console.WriteLine($"TOTAL ERRORS: {n}");
return n > 0 ? 1 : 0;
