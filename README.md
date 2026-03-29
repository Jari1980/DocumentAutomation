# DocumentAutomation

A .NET 10 console app to generate OpenOffice/LibreOffice text documents (`.odt`) or Word documents (`.docx`) from templates with placeholders.

## Features

- Uses template files (`.ott` or `.dotx`) with placeholders like `{{Date}}`, `{{CustomerName}}`, `{{Item1}}`, etc.
- Generates documents automatically into an `Output/` folder.
- Template can include tables, headers, images, and formatted text.
- Supports multiple templates without extra code.
- Minimal dependencies outside .NET 10 and Open XML SDK.

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0)
- Visual Studio 2022 Community (or any IDE that supports .NET 10)
- OpenOffice 4.x or LibreOffice to edit `.ott` templates
- NuGet package: `DocumentFormat.OpenXml`  

# Usage

1, Place your template (.ott or .dotx) in the Templates/ folder.

2, Update the Program.cs file to provide replacement values:
```bash
var replacements = new Dictionary<string, string>
{
    { "{{Date}}", DateTime.Now.ToShortDateString() },
    { "{{CustomerName}}", "John Doe" },
    { "{{Item1}}", "Apples" },
    { "{{Qty1}}", "5" },
    { "{{Item2}}", "Bananas" },
    { "{{Qty2}}", "3" }
};
```
3, Run the project. The output file will be generated in Output/.
