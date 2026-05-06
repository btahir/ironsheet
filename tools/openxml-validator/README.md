# Open XML SDK Validator

This optional harness validates XLSX/XLSM packages with the official .NET Open XML SDK.

Run it directly when the .NET SDK is installed:

```bash
dotnet run --project tools/openxml-validator/OpenXmlValidator.csproj -- path/to/workbook.xlsx
```

Run it through the compatibility checker:

```bash
IRONSHEET_RUN_OPENXML_SDK=1 npm run compat:check -- path/to/workbook.xlsx
```

The local CI gate does not require .NET. When .NET is unavailable, compatibility reports mark this validator as skipped.
