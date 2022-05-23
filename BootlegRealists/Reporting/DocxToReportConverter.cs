using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting;

/// <summary>
/// Base class for the docx to report converter.
/// </summary>
public abstract class DocxToReportConverter : ReportConverter
{
	/// <summary>
	/// The run properties
	/// </summary>
	IDictionary<string, string>? runProps;

	/// <inheritdoc />
	public override void Execute(Stream inputStream, Stream outputStream,
		IDictionary<string, string>? runProperties = null)
	{
		runProps = runProperties;
		ExecuteCore(inputStream, outputStream);
	}

	/// <summary>
	/// Converts the input stream (containing the source report) to the output stream (which will contain the destination
	/// report).
	/// </summary>
	/// <param name="inputStream">Input stream (source report)</param>
	/// <param name="outputStream">Output stream (destination report)</param>
	protected abstract void ExecuteCore(Stream inputStream, Stream outputStream);

	/// <summary>
	/// Gets the run properties
	/// </summary>
	/// <returns>The run properties</returns>
	protected IDictionary<string, string>? GetRunProperties() => runProps;

	/// <summary>
	/// Sets the field codes to their values for the given report
	/// </summary>
	/// <param name="inputStream">Stream containing the report</param>
	protected void SetFieldCodes(Stream inputStream)
	{
		using var docxDocument = WordprocessingDocument.Open(inputStream, true);
		FieldCodeToSimpleField(docxDocument);

		if (docxDocument.MainDocumentPart?.Document.Body == null) return;
		var fields = docxDocument.MainDocumentPart.Document.Body.Descendants<SimpleField>().ToList();
		fields.AddRange(
			docxDocument.MainDocumentPart.HeaderParts.SelectMany(x => x.Header.Descendants<SimpleField>()));
		fields.AddRange(
			docxDocument.MainDocumentPart.FooterParts.SelectMany(x => x.Footer.Descendants<SimpleField>()));
		foreach (var field in fields)
		{
			var s = SimpleFieldToValue(field);
			field.Parent?.ReplaceChild(new Run(new Text(s)), field);
		}

		var runs = docxDocument.MainDocumentPart.Document.Body.Descendants<Run>().ToList();
		runs.AddRange(docxDocument.MainDocumentPart.HeaderParts.SelectMany(x => x.Header.Descendants<Run>()));
		runs.AddRange(docxDocument.MainDocumentPart.FooterParts.SelectMany(x => x.Footer.Descendants<Run>()));
		foreach (var run in runs) RunFieldCodeToValue(run);
	}

	/// <summary>
	/// Converts field codes to simple fields for the given document.
	/// </summary>
	/// <param name="document">Document to process</param>
	protected static void FieldCodeToSimpleField(WordprocessingDocument document)
	{
		if (document.MainDocumentPart?.Document.Body == null) return;
		var paragraphs = document.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();
		paragraphs.AddRange(
			document.MainDocumentPart.HeaderParts.SelectMany(x => x.Header.Descendants<Paragraph>()));
		paragraphs.AddRange(
			document.MainDocumentPart.FooterParts.SelectMany(x => x.Footer.Descendants<Paragraph>()));
		foreach (var paragraph in paragraphs) FieldCodeToSimpleField(paragraph);
	}

	/// <summary>
	/// Converts the given simple field to its value
	/// </summary>
	/// <param name="simpleField">Simple field to process</param>
	/// <returns>The value or "" otherwise</returns>
	protected string SimpleFieldToValue(SimpleField simpleField)
	{
		if (simpleField.Instruction == null || simpleField.Instruction.Value == null) return "";
		var instruction = simpleField.Instruction.Value;
		var instructionItems = instruction.Split(' ', StringSplitOptions.RemoveEmptyEntries);
		if (instructionItems.Length == 0) return "";
		var command = instructionItems[0];

		// TODO: \* MERGEFORMAT, \@ "HH:mm:ss", \# etc...

		var fieldCodes = new Dictionary<string, Func<string[], string>>
		{
			{
				"DATE", _ =>
				{
					var now = DateTime.Now;
					return $"{now.ToShortDateString()} {now.ToShortTimeString()}";
				}
			},
			{
				"TIME", _ =>
				{
					var now = DateTime.Now;
					return $"{now.ToShortDateString()} {now.ToShortTimeString()}";
				}
			},
			{"USERNAME", _ => GetDocumentProperty("UserName")},
			{"TITLE", _ => GetDocumentProperty("Title")},
			{"DOCVARIABLE", GetDocumentVariableProperty}
		};

		string? s = null;
		if (fieldCodes.TryGetValue(command, out var func))
			s = func(instructionItems);

		return !string.IsNullOrEmpty(s) ? s : "";
	}

	/// <summary>
	/// Converts a run with field codes to its values
	/// </summary>
	/// <param name="run">Run to convert</param>
	void RunFieldCodeToValue(OpenXmlElement run)
	{
		var fieldCodes = new Dictionary<string, Func<string[], string>>
		{
			{
				"%DateTime%", _ =>
				{
					var now = DateTime.Now;
					return $"{now.ToShortDateString()} {now.ToShortTimeString()}";
				}
			},
			{"%UserName%", _ => GetDocumentProperty("UserName")},
			{"%Title%", _ => GetDocumentProperty("Title")},
			{"%ComputerName%", _ => Environment.MachineName}
		};

		foreach (var (key, value) in fieldCodes)
		{
			var replace = key;
			foreach (var t in run.Descendants<Text>()
				         .Where(x => !string.IsNullOrEmpty(x.Text) && x.Text.Contains(replace)))
				t.Text = t.Text.Replace(replace, value(Array.Empty<string>()));
		}
	}

	/// <summary>
	/// Converts field codes to simple fields for the given element.
	/// </summary>
	/// <param name="mainElement">Element to process</param>
	static void FieldCodeToSimpleField(OpenXmlElement mainElement)
	{
		//  search for all the Run elements
		var runs = mainElement.Descendants<Run>().ToArray();
		if (runs.Length == 0) return;

		var newFields = new Dictionary<Run, Run[]>();
		var cursor = 0;
		do
		{
			var run = runs[cursor];

			var fieldChar = run.Descendants<FieldChar>().FirstOrDefault(x =>
				x.FieldCharType != null && x.FieldCharType == FieldCharValues.Begin);
			if (fieldChar != null)
			{
				var innerRuns = new List<Run> { run };

				//  loop until we find the 'end' FieldChar
				bool endChar;
				string? instruction = null;
				RunProperties? runProp;
				do
				{
					cursor++;
					run = runs[cursor];

					innerRuns.Add(run);
					var fieldCode = run.Descendants<FieldCode>().FirstOrDefault();
					if (fieldCode != null)
						instruction += fieldCode.Text;
					fieldChar = run.Descendants<FieldChar>().FirstOrDefault(x =>
						x.FieldCharType != null && x.FieldCharType == FieldCharValues.End);
					endChar = fieldChar != null;
					runProp = run.Descendants<RunProperties>().FirstOrDefault();
				} while (!endChar && cursor < runs.Length - 1);

				if (string.IsNullOrEmpty(instruction)) continue;

				var newRun = new Run();
				// must be before simple field set all properties
				if (runProp != null)
					newRun.AppendChild(runProp.CloneNode(true));
				// no field date format
				var simpleField = new SimpleField
				{
					Instruction = instruction
				};
				newRun.AppendChild(simpleField);
				newFields.Add(newRun, innerRuns.ToArray());
			}

			cursor++;
		} while (cursor < runs.Length);

		//  replace all FieldCodes by old-style SimpleFields
		foreach (var (key, value) in newFields)
		{
			value[0].Parent?.ReplaceChild(key, value[0]);
			for (var i = 1; i < value.Length; i++)
				value[i].Remove();
		}
	}

	/// <summary>
	/// Gets a document run property value by name. For example: "userName" -> "james.brighton@protonmail.com"
	/// </summary>
	/// <param name="name">Name of the property</param>
	/// <returns>The value or null otherwise</returns>
	string GetDocumentProperty(string name)
	{
		if (runProps == null || string.IsNullOrEmpty(name)) return "";
		return runProps.TryGetValue(name, out var v) ? v : "";
	}

	/// <summary>
	/// Gets a DOCVARIABLE field from the document run properties.
	/// </summary>
	/// <param name="instructionItems">Items in the instruction</param>
	/// <returns>The value or "" otherwise</returns>
	string GetDocumentVariableProperty(string[] instructionItems)
	{
		if (runProps == null || instructionItems.Length < 2) return "";
		var name = instructionItems[1];

		var dict = new Dictionary<string, Func<string>>
		{
			{"ComputerName", () => Environment.MachineName}
		};

		string? s = null;
		if (dict.TryGetValue(name, out var func))
			s = func();

		return s ?? "";
	}
}
