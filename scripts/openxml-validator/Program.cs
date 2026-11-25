// Validates Word and Excel packages with Microsoft's Open XML SDK (`OpenXmlValidator`), and prints one JSON object
// per file: its path and every error the validator reports, or why it could not be read. Built and run by
// `scripts/verify-openxml.ts`.
//
// Usage: dotnet run -- <file>...    (or file paths on stdin, one per line, when none are given)

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var files = args.Length > 0 ? args : ReadLines();
var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365) { MaxNumberOfErrors = 1000 };

foreach (var file in files)
{
    object result;
    try
    {
        using var package = Open(file);
        var errors = validator.Validate(package).Select(error => new
        {
            part = error.Part?.Uri.ToString(),
            path = error.Path?.XPath,
            id = error.Id,
            type = error.ErrorType.ToString(),
            description = error.Description,
            // The element the report is about, as written — what a known disagreement is checked against, since the
            // description alone does not say what value an attribute had.
            node = Truncate(error.Node?.OuterXml),
        });
        result = new { file, errors = errors.ToList() };
    }
    catch (Exception exception)
    {
        result = new { file, failure = $"{exception.GetType().Name}: {exception.Message}" };
    }
    Console.WriteLine(JsonSerializer.Serialize(result));
}

static OpenXmlPackage Open(string file) => Path.GetExtension(file).ToLowerInvariant() switch
{
    ".docx" or ".docm" or ".dotx" or ".dotm" => WordprocessingDocument.Open(file, false),
    ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => SpreadsheetDocument.Open(file, false),
    var other => throw new NotSupportedException($"{other} is not an Open XML package this validator reads"),
};

static IEnumerable<string> ReadLines()
{
    string? line;
    while ((line = Console.ReadLine()) != null)
    {
        if (line.Length > 0)
        {
            yield return line;
        }
    }
}

static string? Truncate(string? xml) => xml is null || xml.Length <= 2000 ? xml : xml[..2000];
