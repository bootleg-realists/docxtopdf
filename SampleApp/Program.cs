using System.Globalization;
using BootlegRealists.Reporting;

var docxFileName = Array.Find(args, x => x.ToLower(CultureInfo.InvariantCulture).EndsWith(".docx"));
var pdfFileName = Array.Find(args, x => x.ToLower(CultureInfo.InvariantCulture).EndsWith(".pdf"));
if (string.IsNullOrEmpty(docxFileName) || string.IsNullOrEmpty(pdfFileName))
	return;

using var docxStream = new FileStream(docxFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
using var pdfStream = new FileStream(pdfFileName, FileMode.Create, FileAccess.Write, FileShare.Write);
var docxToPdf = new DocxToPdf();
var runProperties = new Dictionary<string, string> { ["Title"] = "title", ["UserName"] = "userName" };
docxToPdf.Execute(docxStream, pdfStream, runProperties);
