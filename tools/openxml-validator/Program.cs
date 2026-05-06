using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length != 1)
{
    Console.Error.WriteLine("usage: dotnet run --project tools/openxml-validator/OpenXmlValidator.csproj -- path/to/workbook.xlsx");
    return 2;
}

var workbookPath = Path.GetFullPath(args[0]);
if (!File.Exists(workbookPath))
{
    Console.Error.WriteLine($"Workbook file does not exist: {workbookPath}");
    return 2;
}

try
{
    using var document = SpreadsheetDocument.Open(workbookPath, false);
    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
    var errors = validator.Validate(document).ToArray();
    var report = new OpenXmlValidationReport(
        SchemaVersion: 1,
        WorkbookPath: workbookPath,
        ErrorCount: errors.Length,
        Errors: errors.Select(OpenXmlValidationError.From).ToArray()
    );

    Console.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    return errors.Length == 0 ? 0 : 1;
}
catch (Exception exception)
{
    Console.Error.WriteLine(exception.Message);
    return 1;
}

internal sealed record OpenXmlValidationReport(
    int SchemaVersion,
    string WorkbookPath,
    int ErrorCount,
    OpenXmlValidationError[] Errors
);

internal sealed record OpenXmlValidationError(
    string? Id,
    string? Description,
    string? ErrorType,
    string? PartUri,
    string? Path
)
{
    internal static OpenXmlValidationError From(ValidationErrorInfo error)
    {
        return new OpenXmlValidationError(
            Id: error.Id,
            Description: error.Description,
            ErrorType: error.ErrorType.ToString(),
            PartUri: error.Part?.Uri.ToString(),
            Path: error.Path?.XPath
        );
    }
}
