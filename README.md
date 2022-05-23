# Introduction

DocxToPdf is a class library for converting a Word document (\*.docx) to a PDF document (\*.pdf). Based on the hard work by https://github.com/LINALIN1979/DocxToPdf.

There's still a lot of functionality missing (any help is appreciated). But with a basic document the output isn't half bad. 

## Additions

- .NET Standard 2.1
- Restructured code
- Footer
- Font resolver
- Background color
- Support for macOS

## How to use it

See the sample application "SampleApp" to convert a document.

````csharp
using var docxStream = new FileStream(docxFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
using var pdfStream = new FileStream(pdfFileName, FileMode.Create, FileAccess.Write, FileShare.Write);
var docxToPdf = new DocxToPdf();
var runProperties = new Dictionary<string, string> { ["Title"] = "title", ["UserName"] = "userName" };
docxToPdf.Execute(docxStream, pdfStream, runProperties);
````