using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using BootlegRealists.Reporting.Extension;
using BootlegRealists.Reporting.Function;
using Pdf = iTextSharp.text.pdf;
using Text = iTextSharp.text;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class converts Word documents to PDF files.
/// </summary>
public partial class DocxToPdf : DocxToReportConverter
{
	/// <summary>
	/// Initializes any static data of the <see cref="DocxToPdf" /> class.
	/// </summary>
	static DocxToPdf() => AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);

	CounterHelper counterHelper = new();
	WordprocessingDocument docxDocument = WordprocessingDocumentEx.CreateEmpty();
	Text.DocumentEx pdfDocument = new();

	/// <inheritdoc />
	protected override void ExecuteCore(Stream inputStream, Stream outputStream)
	{
		// Create a copy as it will be modified
		using var memoryStream = new MemoryStream();
		inputStream.CopyTo(memoryStream);
		memoryStream.Position = 0;
		SetFieldCodes(memoryStream);

		using (docxDocument = WordprocessingDocument.Open(memoryStream, true))
		{
			FieldCodeToSimpleField(docxDocument);
			counterHelper = new CounterHelper(docxDocument);

			var body = docxDocument.MainDocumentPart?.Document.Body;
			if (body == null) return;
			CreatePdfDocument(body);
			var writer = Pdf.PdfWriterEx.GetInstance(pdfDocument, outputStream);
			var pdfPageEvent = new PdfPageEvent(this);
			writer.PageEvent = pdfPageEvent;
			pdfDocument.Open();

			var list = body.Elements().ToList();
			Text.Paragraph? previousParagraph = null;
			var elements = new List<Text.IElement>();
			for (var i = 0; i < list.Count; i++)
			{
				var element = Dispatcher(list, i, previousParagraph);
				previousParagraph = element as Text.Paragraph;
				if (element == null) continue;
				elements.Add(element);
			}

			foreach (var element in pdfDocument.Process(elements))
				pdfDocument.Add(element);
			pdfDocument.Close();
		}
	}

	/// <summary>
	/// Creates the PDF document
	/// </summary>
	/// <param name="body">Body to use for the document</param>
	void CreatePdfDocument(OpenXmlElement body)
	{
		var section = body.Descendants<SectionProperties>().FirstOrDefault();
		var size = section?.Descendants<PageSize>().FirstOrDefault();
		var margin = section?.Descendants<PageMargin>().FirstOrDefault();
		var pageSize = new Text.Rectangle(Converter.TwipToPoint(size?.Width?.Value ?? 0.0f),
			Converter.TwipToPoint(size?.Height?.Value ?? 0.0f));
		var leftMargin = Converter.TwipToPoint(margin?.Left?.Value ?? 0.0f);
		var rightMargin = Converter.TwipToPoint(margin?.Right?.Value ?? 0.0f);
		var topMargin = Converter.TwipToPoint(margin?.Top?.Value ?? 0.0f);
		var bottomMargin = Converter.TwipToPoint(margin?.Bottom?.Value ?? 0.0f);

		pdfDocument = new Text.DocumentEx(pageSize, leftMargin, rightMargin, topMargin, bottomMargin);
	}

	/// <summary>
	/// Gets the shading color for the given open xml element
	/// </summary>
	/// <param name="element">Given element</param>
	/// <returns>The color or null otherwise</returns>
	static Text.BaseColor? GetShadingColor(OpenXmlElement element)
	{
		// Background/shading
		var shading = element.GetEffectiveElement<Shading>();
		if (shading?.Fill?.HasValue == true && shading.Fill.Value != "auto")
			return new Text.BaseColor(Convert.ToInt32(shading.Fill.Value, 16));
		return null;
	}

	/// <summary>
	/// Handle container (i.e. table and paragraph)
	/// </summary>
	/// <param name="list">List with elements.</param>
	/// <param name="index">Index to process.</param>
	/// <param name="previous">Previous paragraph (can be null).</param>
	/// <returns>The converted element</returns>
	Text.IElement? Dispatcher(IList<OpenXmlElement> list, int index, Text.Paragraph? previous)
	{
		return list[index] switch
		{
			Paragraph => BuildParagraph(list, index, previous),
			Table table => BuildTable(table),
			_ => null
		};
	}

	/// <summary>
	/// Set paragraph's linespacing, spacingBefore and spacingAfter.
	/// </summary>
	/// <param name="list">List containing the element to process</param>
	/// <param name="index">Index of element in list to process</param>
	/// <param name="pgPrevious">Previous paragraph</param>
	/// <param name="pg">Destination paragraph</param>
	static void SetParagraphSpacing(IList<OpenXmlElement> list, int index, Text.Paragraph? pgPrevious, Text.Paragraph pg)
	{
		const float wordDefaultLineSpacing = 1.15f; // magic: word default leading

		if (list[index] is not Paragraph paragraph) return;

		var contextualSpacing = paragraph.UsesContextualSpacing();
		var previousParagraph = index > 0 ? list[index - 1] as Paragraph : null;
		var nextParagraph = index < list.Count - 1 ? list[index + 1] as Paragraph : null;
		var previousStyleSame = previousParagraph != null && paragraph.HasSameStyle(previousParagraph);
		var nextStyleSame = nextParagraph != null && paragraph.HasSameStyle(nextParagraph);

		float spacingAfter = float.NaN, spacingBefore = float.NaN, linespacing = float.NaN;
		var space = paragraph.GetEffectiveElement<SpacingBetweenLines>();
		var paragraphFontSize = pg.GetCalculatedFont()?.CalculatedSize ?? 16.0f;

		if (space?.LineRule != null && space.Line != null && space.LineRule.HasValue && space.Line.HasValue)
		{
			if (space.LineRule.Value == LineSpacingRuleValues.AtLeast) // interpreted as twip
			{			
				var spacePoint = Converter.TwipToPoint(space.Line?.Value ?? "");
				if (spacePoint >= paragraphFontSize * wordDefaultLineSpacing)
					linespacing = spacePoint;
				else
					linespacing = wordDefaultLineSpacing * paragraphFontSize;
			}
			else if (space.LineRule.Value == LineSpacingRuleValues.Exact) // interpreted as twip
			{
				linespacing = Converter.TwipToPoint(space.Line?.Value ?? "");
			}
			else if (space.LineRule.Value == LineSpacingRuleValues.Auto) // interpreted as 240th of a line
			{
				linespacing = Convert.ToSingle(space.Line.Value, CultureInfo.InvariantCulture) / 240 * paragraphFontSize * wordDefaultLineSpacing;
			}
		}

		if (space?.After?.HasValue == true)
			spacingAfter = Converter.TwipToPoint(space.After?.Value ?? "");
		else if (space?.AfterLines?.HasValue == true)
			spacingAfter = Convert.ToSingle(space.After?.Value ?? "", CultureInfo.InvariantCulture) / 100 * paragraphFontSize;

		if (space?.Before?.HasValue == true)
			spacingBefore = Converter.TwipToPoint(space.Before?.Value ?? "");
		else if (space?.BeforeLines?.HasValue == true)
			spacingBefore = Convert.ToSingle(space.Before?.Value ?? "", CultureInfo.InvariantCulture) / 100 * paragraphFontSize;

		if (float.IsNaN(linespacing))
			linespacing = wordDefaultLineSpacing * paragraphFontSize;

		pg.SetLeading(linespacing, 0f);

		if (!float.IsNaN(spacingAfter) && (!contextualSpacing || !nextStyleSame)) pg.SpacingAfter = spacingAfter;

		if (float.IsNaN(spacingBefore) || !contextualSpacing && previousStyleSame)
			return;

		pg.SpacingBefore = spacingBefore;
		if (pgPrevious == null) return;
		// The bigger one of two paragraphs is used
		if (pgPrevious.SpacingAfter > pg.SpacingBefore)
			pg.SpacingBefore = 0.0f;
		else
			pgPrevious.SpacingAfter = 0.0f;
	}

	/// <summary>
	/// Sets the horizontal justification
	/// </summary>
	/// <param name="paragraph">Source paragraph</param>
	/// <param name="pg">Destination paragraph</param>
	static void SetHorizontalJustification(OpenXmlElement paragraph, Text.Paragraph pg)
	{
		// Horizontal Justification
		var jc = paragraph.GetEffectiveElement<Justification>();
		if (jc?.Val == null) return;
		if (jc.Val.Value == JustificationValues.Center)
		{
			pg.Alignment = Text.Element.ALIGN_CENTER;
		}
		else if (jc.Val.Value == JustificationValues.Left)
		{
			pg.Alignment = Text.Element.ALIGN_LEFT;
		}
		else if (jc.Val.Value == JustificationValues.Right)
		{
			pg.Alignment = Text.Element.ALIGN_RIGHT;
		}	
		else if (jc.Val.Value == JustificationValues.Both|| jc.Val.Value == JustificationValues.Distribute)
		{
			// justify text between both margins equally, and both inter-word and inter-character spacing are affected. iTextSharp doesnt support this.
			pg.Alignment = Text.Element.ALIGN_JUSTIFIED;
		}
	}

	/// <summary>
	/// Handle autoSpaceDE and autoSpaceDN of paragraph.
	/// autoSpaceDE: Automatically Adjust Spacing of Latin and East Asian Text
	/// autoSpaceDN: Automatically Adjust Spacing of Number and East Asian Text
	/// </summary>
	/// <param name="paragraph">Source paragraph</param>
	/// <param name="pg">Destination paragraph</param>
	static void SetParagraphAutoSpace(OpenXmlElement paragraph, Text.Phrase pg)
	{
		var asde = paragraph.GetEffectiveElement<AutoSpaceDE>();
		var autoSpaceDe = asde == null || asde.Val?.Value != false; // omitted means true, different from other toggle property

		var asdn = paragraph.GetEffectiveElement<AutoSpaceDN>();
		var autoSpaceDn = asdn == null || asdn.Val?.Value != false; // omitted means true, different from other toggle property

		if (!autoSpaceDe && !autoSpaceDn)
			return;

		if (pg.Chunks.Count < 2)
			return;

		var chunks = pg.Chunks.Cast<Text.Chunk>().ToList();
		pg.Clear();
		pg.AddRange(chunks);
		for (int i = pg.Chunks.Count - 2, j = i + 1; i >= 0; i--, j = i + 1)
		{
			if (pg.Chunks[i] is not Text.Chunk ich || pg.Chunks[j] is not Text.Chunk jch) continue;

			// bypass line break & page break
			if (string.Equals(ich.Content, ChunkFunction.NewLine.Content, StringComparison.Ordinal) || string.Equals(ich.Content, ChunkFunction.NextPage.Content, StringComparison.Ordinal) || string.Equals(jch.Content, ChunkFunction.NewLine.Content, StringComparison.Ordinal) || string.Equals(jch.Content, ChunkFunction.NextPage.Content, StringComparison.Ordinal))
				continue;

			var frontChar = ich.Content[^1];
			var rearChar = jch.Content[0];
			var frontCharType = CodePointRecognizer.GetFontType(frontChar);
			var rearCharType = CodePointRecognizer.GetFontType(rearChar);

			// bypass space and line feed
			var ignoredChars = new List<char> { ' ', '\u00A0', '\n' };
			if (ignoredChars.Contains(frontChar) || ignoredChars.Contains(rearChar)) continue;

			if ((frontCharType.FontType != FontTypeEnum.EastAsian || rearCharType.FontType is FontTypeEnum.EastAsian or FontTypeEnum.ComplexScript) && (rearCharType.FontType != FontTypeEnum.EastAsian || frontCharType.FontType is FontTypeEnum.EastAsian or FontTypeEnum.ComplexScript))
				continue;

			if ((!autoSpaceDn || !char.IsNumber(frontChar) && !char.IsNumber(rearChar)) && (!autoSpaceDe || !char.IsLetter(frontChar) && !char.IsLetter(rearChar)))
				continue;

			//Text.Chunk space = new Text.Chunk('\u00A0');
			var space = new Text.Chunk(' ')
			{
				// Due to we multiply 0.875 to font size, the other font styles 
				// (e.g. underline) are scaled as well, so do not duplicate 
				// font style settings
				//space.Font = new Text.Font(pg.Chunks[i].Font); 
				//space.Font.SetStyle(pg.Chunks[i].Font.Style);
				Font = { Size = (float)(ich.Font.CalculatedSize * 0.875) }
			};
			pg.Insert(j, space);
		}
	}

	/// <summary>
	/// Handle paragraph indentation and numbering/listing.
	/// </summary>
	/// <param name="paragraph">Source paragraph</param>
	/// <param name="pg">Destination paragraph</param>
	void SetParagraphIndentation(Paragraph paragraph, Text.Paragraph pg)
	{
		var level = counterHelper.GetLevel(paragraph);

		// Generate numbering/listing text
		// https://msdn.microsoft.com/en-us/library/office/ee922775%28v=office.14%29.aspx
		Text.Chunk? numbering = null;
		if (level != null)
		{
			var text = GenerateNumbering(level);
			var font = GetNumberingFont(paragraph, level, text);
			numbering = new Text.Chunk(text, font);
		}

		// Indentation apply order: direct formatting > numbering > paragraph property
		var dirind = paragraph.ParagraphProperties?.GetFirstDescendant<Indentation>();
		var numind = level?.PreviousParagraphProperties?.Indentation;
		var pgind = dirind == null ? paragraph.GetEffectiveElement<Indentation>() : null;
		var ind = new Indentation();
		pgind?.AppendAttributesTo(ind);
		numind?.AppendAttributesTo(ind);

		dirind?.AppendAttributesTo(ind);

		// Character Unit of hanging and firstLine indentation are
		// based on the font size of the first character of paragraph
		// source: https://social.msdn.microsoft.com/Forums/office/en-US/3cfbd59e-453d-4d7e-9bc8-ecb417dbe4a7/how-many-twips-is-a-character-unit?forum=oxmlsdk

		// hanging, use the first character's font size for HaningChars
		var hanging = float.NaN;
		if (ind.Hanging?.HasValue == true)
		{
			hanging = Converter.TwipToPoint(ind.Hanging?.Value);
		}
		else if (ind.HangingChars?.HasValue == true)
		{
			var firstChunk = pg.Count > 0 ? pg.Chunks[0] as Text.Chunk : null;
			hanging = Converter.HundredthOfCharacterToPoint(ind.HangingChars.Value,
				firstChunk?.Font.CalculatedSize ?? 0.0f);
		}

		// firstLine (only available when no hanging), use the first character's font size for FirstLineChars
		var firstline = -1f;
		if (float.IsNaN(hanging))
		{
			if (ind.FirstLine?.HasValue == true)
			{
				firstline = Converter.TwipToPoint(ind.FirstLine?.Value);
			}
			else if (ind.FirstLineChars?.HasValue == true)
			{
				var firstChunk = pg.Count > 0 ? pg.Chunks[0] as Text.Chunk : null;
				firstline = Converter.HundredthOfCharacterToPoint(ind.FirstLineChars.Value,
					firstChunk?.Font.CalculatedSize ?? 0.0f);
			}
		}

		// Character Unit of start and end are based on the font size
		// of paragraph style hierarchy
		var fontSizeInPoints = 12f;
		var fontSize = paragraph.GetEffectiveElement<FontSize>();
		if (fontSize?.Val != null)
			fontSizeInPoints = Converter.HalfPointToPoint(fontSize.Val?.Value);

		// start
		var dist = 0f;
		if (ind.Left?.HasValue == true)
			dist = Converter.TwipToPoint(ind.Left?.Value);
		else if (ind.Start?.HasValue == true)
			dist = Converter.TwipToPoint(ind.Start?.Value);
		else if (ind.LeftChars?.HasValue == true)
			dist = Converter.HundredthOfCharacterToPoint(ind.LeftChars.Value, fontSizeInPoints);
		else if (ind.StartCharacters?.HasValue == true)
			dist = Converter.HundredthOfCharacterToPoint(ind.StartCharacters.Value, fontSizeInPoints);

		if (hanging >= 0f)
		{
			// first line indentation is based on IndentationLeft to add/reduce
			pg.IndentationLeft = dist;
			pg.FirstLineIndent = -hanging;
		}
		else
		{
			// first line indentation is based on IndentationLeft to add/reduce
			pg.IndentationLeft = dist;
			if (firstline >= 0f) pg.FirstLineIndent = firstline;
		}

		// end
		dist = 0f;
		if (ind.Right?.HasValue == true)
			dist = Converter.TwipToPoint(ind.Right?.Value);
		else if (ind.End?.HasValue == true)
			dist = Converter.TwipToPoint(ind.End?.Value);
		else if (ind.RightChars?.HasValue == true)
			dist = Converter.HundredthOfCharacterToPoint(ind.RightChars.Value, fontSizeInPoints);
		else if (ind.EndCharacters?.HasValue == true)
			dist = Converter.HundredthOfCharacterToPoint(ind.EndCharacters.Value, fontSizeInPoints);

		if (dist != 0f)
			pg.IndentationRight = dist;

		if (numbering == null)
			return;

		var addNumberingSpaceWidth = pg.IndentationLeft - (pg.IndentationLeft + pg.FirstLineIndent + numbering.GetWidthPoint());
		if (addNumberingSpaceWidth > 0f)
		{
			var space = new Text.SpaceChunk("\u0020", new Text.Font(numbering.Font));
			space.Font.Size = numbering.Font.CalculatedSize * (addNumberingSpaceWidth / space.GetWidthPoint());
			pg.Insert(0, space);
		}

		pg.Insert(0, numbering);
	}

	Text.Font? GetNumberingFont(Paragraph paragraph, Level level, string text)
	{
		if (string.IsNullOrEmpty(text)) return null;
		// Get bullet font's RunFonts and size
		var fontType = CodePointRecognizer.GetFontType(text[0]);
		BaseFontEx? baseFont = null;
		FontSizeComplexScript? fscs = null;
		FontSize? fs = null;
		//  get from numbering rPr
		var bulletRunFonts = level.NumberingSymbolRunProperties?.RunFonts;
		if (bulletRunFonts != null)
		{
			baseFont = FontCreator.GetBaseFontByFontName(docxDocument, FontCreator.GetFontNameFromRunFontsByFontType(bulletRunFonts, fontType),
				false, false,
				fontType,
				level.NumberingSymbolRunProperties?.GetEffectiveElement<Languages>());
			if (fontType.FontType == FontTypeEnum.ComplexScript)
				fscs = level.GetFirstDescendant<FontSizeComplexScript>();
			else
				fs = level.GetFirstDescendant<FontSize>();
		}

		//  get from paragraph rPr (i.e. RunFonts for paragraph glyph)
		if (baseFont == null && paragraph.ParagraphProperties?.ParagraphMarkRunProperties != null)
		{
			bulletRunFonts = paragraph.ParagraphProperties.ParagraphMarkRunProperties
				.Descendants<RunFonts>().FirstOrDefault();
			if (bulletRunFonts != null)
			{
				baseFont = FontCreator.GetBaseFontByFontName(docxDocument, 
					FontCreator.GetFontNameFromRunFontsByFontType(bulletRunFonts, fontType),
					false,
					false,
					fontType,
					paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetEffectiveElement<Languages>());
				if (fontType.FontType == FontTypeEnum.ComplexScript)
					fscs = paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstDescendant<FontSizeComplexScript>();
				else
					fs = paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstDescendant<FontSize>();
			}
		}

		// StyleId
		if (baseFont == null && paragraph.ParagraphProperties?.ParagraphStyleId != null)
		{
			var style = paragraph.ParagraphProperties?.ParagraphStyleId.GetStyleById()!;
			var bulletRunFonts2 = style.GetEffectiveElement<RunFonts>();
			if (bulletRunFonts2 != null)
			{
				baseFont = FontCreator.GetBaseFontByFontName(docxDocument, 
					FontCreator.GetFontNameFromRunFontsByFontType(bulletRunFonts2, fontType),
					false,
					false,
					fontType,
					paragraph.GetEffectiveElement<Languages>());
				if (fontType.FontType == FontTypeEnum.ComplexScript)
					fscs = paragraph.GetEffectiveElement<FontSizeComplexScript>();
				else
					fs = paragraph.GetEffectiveElement<FontSize>();
			}
		}

		//  get from docDefault rPr
		if (baseFont == null)
		{
			bulletRunFonts = docxDocument.MainDocumentPart?.GetDocDefaults<RunFonts>(DocDefaultsType.Character);
			if (bulletRunFonts != null)
			{
				baseFont = FontCreator.GetBaseFontByFontName(docxDocument, 
					FontCreator.GetFontNameFromRunFontsByFontType(bulletRunFonts, fontType),
					false,
					false,
					fontType,
					docxDocument.MainDocumentPart?.GetDocDefaults<Languages>(DocDefaultsType.Character));
				if (fontType.FontType == FontTypeEnum.ComplexScript)
					fscs = docxDocument.MainDocumentPart?.GetDocDefaults<FontSizeComplexScript>(DocDefaultsType.Character);
				else
					fs = docxDocument.MainDocumentPart?.GetDocDefaults<FontSize>(DocDefaultsType.Character);
			}
		}

		Text.Font? font = null;
		if (baseFont == null) return font;
		var f = 12f;
		if (fscs?.Val != null)
			f = Converter.HalfPointToPoint(fscs.Val?.Value);
		else if (fs?.Val != null)
			f = Converter.HalfPointToPoint(fs.Val?.Value);

		font = FontFactory.CreateFont(baseFont, f,
			bulletRunFonts?.Descendants<Bold>().FirstOrDefault(),
			bulletRunFonts?.Descendants<Italic>().FirstOrDefault(),
			bulletRunFonts?.Descendants<Strike>().FirstOrDefault(),
			bulletRunFonts?.Descendants<Color>().FirstOrDefault());
		return font;
	}

	string GenerateNumbering(Level level)
	{
		var text = "";
		var lvlText = level.Descendants<LevelText>().FirstOrDefault();
		if (lvlText?.Val != null)
			text = lvlText.Val?.Value ?? "";

		var numFmt = level.Descendants<NumberingFormat>().FirstOrDefault();
		if (numFmt?.Val == null || numFmt.Val.Value == NumberFormatValues.Bullet) return text;
		List<int>? current = null;
		if (level.Parent is AbstractNum an && an.AbstractNumberId != (object?)null &&
		    level.LevelIndex != (object?)null)
			current = counterHelper.GetCurrent(an.AbstractNumberId.Value, level.LevelIndex.Value);

		if (current == null) return "";

		for (var i = 0; i < current.Count; i++)
		{
			var replacePattern = $"%{i + 1}";
			string? str;
			if (numFmt.Val.Value == NumberFormatValues.TaiwaneseCountingThousand)
				str = Tools.IntToTaiwanese(current[i]);
			else if (numFmt.Val.Value == NumberFormatValues.LowerRoman)
				str = Tools.IntToRoman(current[i], false);
			else if (numFmt.Val.Value == NumberFormatValues.UpperRoman)
				str = Tools.IntToRoman(current[i], true);
			else if (numFmt.Val.Value == NumberFormatValues.DecimalZero)
				str = $"0{current[i]}";
			else
				str = current[i].ToString(CultureInfo.InvariantCulture);

			text = text.Replace(replacePattern, str);
		}

		return text;
	}

	/// <summary>
	/// Builds the PDF paragraph from the given OpenXML paragraph (list + index)  
	/// </summary>
	/// <param name="list">List with OpenXML elements</param>
	/// <param name="index">Index in the list with the paragraph</param>
	/// <param name="previous">Previous paragraph</param>
	/// <returns>The created paragraph (PDF) or null otherwise</returns>
	Text.IElement? BuildParagraph(IList<OpenXmlElement> list, int index, Text.Paragraph? previous)
	{
		if (list[index] is not Paragraph paragraph)
			return null;

		var result = new Text.ParagraphEx();
		ProcessElements(paragraph, result);
		SetHorizontalJustification(paragraph, result);

		// w:pPr w:keepLines: all lines of this paragraph are maintained on a single page whenever possible
		result.KeepTogether = Converter.OnOffToBool(paragraph.GetEffectiveElement<KeepLines>());

		// Handle empty paragraph
		if (result.Chunks.Count == 0)
		{
			if (!Converter.OnOffToBool(paragraph.GetEffectiveElement<Vanish>()))
			{
				var emptyPg = new Text.Chunk(" ");
				var fontSize = paragraph.GetEffectiveElement<FontSize>();
				if (fontSize?.Val != null) emptyPg.Font.Size = Converter.HalfPointToPoint(fontSize.Val.Value);

				result.Add(emptyPg);
			}
			else
			{
				return null;
			}
		}

		result.BackgroundColor = GetShadingColor(list[index]);

		SetParagraphAutoSpace(paragraph, result);
		SetParagraphSpacing(list, index, previous, result);
		SetParagraphIndentation(paragraph, result);
		return result;
	}
		
	/// <summary>
	/// Processes the elements in the OpenXML paragraph and adds them (converted) to the result
	/// </summary>
	/// <param name="paragraph">OpenXML paragraph to process elements for</param>
	/// <param name="result">Resulting PDF paragraph to add elements to</param>

	void ProcessElements(OpenXmlElement paragraph, Text.Phrase result)
	{
		foreach (var element in paragraph.Elements())
		{
			switch (element)
			{
				case SimpleField simpleField:
					result.Add(BuildSimpleField(simpleField));
					break;
				case Run run:
					result.AddAll(BuildRun(run));
					break;
				case Hyperlink hyperlink:
					result.Add(BuildHyperlink(hyperlink));
					break;
			}
		}
	}

	/// <summary>
	/// Builds the tab char
	/// </summary>
	/// <param name="tabChar">tab char to process</param>
	/// <returns>The Pdf element or null otherwise</returns>
	Text.IElement? BuildTabChar(OpenXmlElement tabChar)
	{
		var result = new TabCharToPdfElement(docxDocument).Process(tabChar);
		return result.FirstOrDefault();
	}

	/// <summary>
	/// Builds the text
	/// </summary>
	/// <param name="text">Text to process</param>
	/// <returns>The Pdf element or null otherwise</returns>
	Text.IElement? BuildText(OpenXmlElement text)
	{
		var result = new TextTypeToPdfElement(docxDocument).Process(text);
		return result.FirstOrDefault();
	}

	/// <summary>
	/// Builds the simple field.
	/// </summary>
	/// <param name="simpleField">Simple field to process</param>
	/// <returns>The element or null otherwise</returns>
	Text.IElement? BuildSimpleField(SimpleField simpleField)
	{
		var s = SimpleFieldToValue(simpleField);
		if (string.IsNullOrEmpty(s)) return null;

		var ch = new Text.Chunk(s);
		OpenXmlElementFunction.GetCompositeElementEffects(simpleField, out var bold, out var italic, out var strike, out var caps,
			out var uline, out var vertAlignment, out var fontSize, out var csFontSize, out var color);
		return ChunkProcessor.Process(docxDocument, simpleField, ch, bold, italic, strike, caps, uline, vertAlignment, fontSize, csFontSize, color);
	}

	/// <summary>
	/// Builds the hyperlink
	/// </summary>
	/// <param name="hyperlink">Hyperlink to process</param>
	/// <returns>The element or null otherwise</returns>
	Text.IElement? BuildHyperlink(Hyperlink hyperlink)
	{
		var phrase = new Text.Phrase();
		foreach (var element in hyperlink.Elements<Run>()) phrase.AddRange(BuildRun(element));

		if (phrase.Count <= 0)
			return null;

		var anchor = new Text.Anchor(phrase);
		if (hyperlink.Id != null)
			anchor.Reference = hyperlink.GetUrl() + (hyperlink.Anchor?.HasValue == true ? "#" + hyperlink.Anchor.Value : "");

		return anchor;
	}

	/// <summary>
	/// Builds the image from the given picture.
	/// </summary>
	/// <param name="picture">Picture to convert to image.</param>
	/// <returns>The image or null otherwise.</returns>
	Text.Image? BuildPicture(OpenXmlElement picture)
	{
		var result = new PictureToPdfElement(docxDocument).Process(picture);
		return result.FirstOrDefault() as Text.Image;
	}

	/// <summary>
	/// Builds the image from the given drawing.
	/// </summary>
	/// <param name="drawing">Drawing to convert to image.</param>
	/// <returns>The image or null otherwise.</returns>
	Text.Image? BuildDrawing(OpenXmlElement drawing)
	{
		var result = new DrawingToPdfElement(docxDocument).Process(drawing);
		return result.FirstOrDefault() as Text.Image;
	}

	List<Text.Chunk> BuildRun(OpenXmlElement run)
	{
		var ret = new List<Text.Chunk>();

		// Handle RUN properties
		// http://officeopenxml.com/WPtextFormatting.php

		// vanish
		if (Converter.OnOffToBool(run.GetEffectiveElement<Vanish>())) return ret;

		// Toggle property
		OpenXmlElementFunction.GetCompositeElementEffects(run, out var bold, out var italic, out var strike, out var caps, out var uline,
			out var vertAlignment, out var fontSize, out var csFontSize, out var color);

		// Handle OpenXML child elements
		foreach (var element in run.Elements())
		{
			switch (element)
			{
				// Add: TabChar
				case SimpleField s:
				{
					if (BuildSimpleField(s) is not Text.Chunk ch)
						continue;
					ret.Add(ch);
					break;
				}
				case TabChar tabChar:
				{
					if (BuildTabChar(tabChar) is not Text.Chunk ch)
						continue;

					ch = ChunkProcessor.Process(docxDocument, run, ch, bold, italic, strike, caps, uline, vertAlignment, fontSize, csFontSize, color);
					ret.Add(ch);
					break;
				}
				case DocumentFormat.OpenXml.Wordprocessing.Text t:
				{
					if (BuildText(t) is not Text.Chunk ch)
						continue;

					ch = ChunkProcessor.Process(docxDocument, run, ch, bold, italic, strike, caps, uline, vertAlignment, fontSize, csFontSize, color);
					ret.Add(ch);
					break;
				}
				case Break br when br.Type == null || br.Type.Value != BreakValues.Page:
					ret.Add(GetNewLine(run, bold, italic, strike, caps, uline, vertAlignment, fontSize, csFontSize, color));
					continue;
				case Break:
					ret.Add(ChunkFunction.NextPage);
					break;
				case SymbolChar sym:
				{
					var c = (char)(int.Parse(sym.Char?.Value ?? "0x00", NumberStyles.HexNumber, CultureInfo.InvariantCulture) - 0xF000);
					var fontType = CodePointRecognizer.GetFontType(c);
					var baseFont = FontCreator.GetBaseFontByFontName(docxDocument, sym.Font?.Value ?? "", false, false, fontType, null);
					if (baseFont == null)
						continue;

					var ch = new Text.Chunk(c.ToString())
					{
						Font = FontFactory.CreateFont(baseFont, csFontSize > 0 && fontType.FontType == FontTypeEnum.ComplexScript ? csFontSize : fontSize, bold, italic, strike, color)
					};
					ret.Add(ch);
					break;
				}
				case Picture picture:
				{
					var image = BuildPicture(picture);
					if (image != null) ret.Add(new Text.Chunk(image, 0f, 0f, true));
					break;
				}
				case Drawing drawing:
				{
					var image = BuildDrawing(drawing);
					if (image == null)
						continue;

					if (pdfDocument.CurrentHeight > 0.0f && pdfDocument.IndentTop - pdfDocument.CurrentHeight - image.ScaledHeight < pdfDocument.IndentBottom)
						pdfDocument.NewPage();
					ret.Add(new Text.Chunk(image, 0f, 0f, true));
					break;
				}
			}
		}

		return ret;
	}

	/// <summary>
	/// Processes a new line and applies the given effects and font
	/// </summary>
	/// <param name="compositeElement">Composite element referencing the chunk</param>
	/// <param name="bold">Bold effect</param>
	/// <param name="italic">Italic effect</param>
	/// <param name="strike">Strike effect</param>
	/// <param name="caps">Caps effect</param>
	/// <param name="underline">Underline effect</param>
	/// <param name="verticalAlignment">Superscript/subscript effect</param>
	/// <param name="fontSize">Font size (in points)</param>
	/// <param name="fontSizeComplexScript">Complex script font size (in points)</param>
	/// <param name="color">Color</param>
	/// <returns>The processed chunk</returns>
	Text.Chunk GetNewLine(OpenXmlElement compositeElement, Bold? bold, Italic? italic,
		Strike? strike, OnOffType? caps, Underline? underline, VerticalAlignment verticalAlignment, float fontSize, float fontSizeComplexScript,
		Color? color)
	{
		var ch = ChunkProcessor.Process(docxDocument, compositeElement, ChunkFunction.NewLine, bold, italic, strike, caps, underline, verticalAlignment, fontSize, fontSizeComplexScript, color);
		return ch;
	}

	/// <summary>
	/// Convert DocumentFormat.OpenXml.Wordprocessing.Table to iTextSharp.text.pdf.PdfPTable.
	/// </summary>
	/// <param name="table">DocumentFormat.OpenXml.Wordprocessing.Table.</param>
	/// <returns>iTextSharp.text.pdf.PdfPTable or null.</returns>
	Text.IElement BuildTable(Table table)
	{
		var tableHelper = new TableHelper();
		tableHelper.ParseTable(table);

		// ====== Prepare iTextSharp PdfPTable ======

		// Set table width
		var pt = new Pdf.PdfPTable(tableHelper.ColumnLength)
		{
			TotalWidth = tableHelper.TableColumnsWidth?.Sum() ?? 0.0f
		};
		pt.SetWidths(tableHelper.TableColumnsWidth);
		pt.LockedWidth =
			true; // use pt.TotalWidth rather than pt.WidthPercentage (iTextSharp default is WidthPercentage)

		// Table justification
		var jc = table.GetEffectiveElement<TableJustification>();
		pt.HorizontalAlignment = jc?.Val != null
			? jc.Val.Value == TableRowAlignmentValues.Center ? Text.Element.ALIGN_CENTER :
			jc.Val.Value == TableRowAlignmentValues.Left ? Text.Element.ALIGN_LEFT :
			jc.Val.Value == TableRowAlignmentValues.Right ? Text.Element.ALIGN_RIGHT :
			Text.Element.ALIGN_LEFT
			: Text.Element.ALIGN_LEFT;

		foreach (TableHelperCell cellHelper in tableHelper)
		{
			TableBuilder.GetRowHeightPadding(tableHelper, cellHelper.RowId, out var topPadding,
				out var bottomPadding);
			// Row height
			float minRowHeight = float.NaN, exactRowHeight = float.NaN;
			var trHeight = cellHelper.Row?.GetEffectiveElement<TableRowHeight>();
			if (trHeight?.Val != (object?)null)
			{
				var heightPoints = Converter.TwipToPoint(trHeight.Val.Value);
				minRowHeight = heightPoints;

				var hrule = trHeight.GetAttributes().FirstOrDefault(c =>
					string.Equals(c.LocalName, "hRule", StringComparison.OrdinalIgnoreCase));
				if (hrule.Value == "exact") exactRowHeight = heightPoints;
			}

			var cell = new Pdf.PdfPCell
			{
				//cell.UseAscender = false; // remove whitespace on top of each cell even padding&leading set to 0, http://stackoverflow.com/questions/9672046/itextsharp-4-1-6-pdf-table-how-to-remove-whitespace-on-top-of-each-cell-pad
				//cell.UseDescender = true;
				Rowspan = cellHelper.RowSpan,
				Colspan = cellHelper.ColSpan
			}; // composite mode, not text mode
			if (cellHelper.Blank)
			{
				cell.Border = Text.Rectangle.NO_BORDER;
			}
			else if (cellHelper.Cell != null)
			{
				var c = cellHelper.Cell;
				// Cell margins
				var m = TableBuilder.GetMargin<LeftMargin, TableCellLeftMargin>(c);
				cell.PaddingLeft = !float.IsNaN(m) ? m : 0.0f;
				m = TableBuilder.GetMargin<RightMargin, TableCellRightMargin>(c);
				cell.PaddingRight = !float.IsNaN(m) ? m : 0.0f;

				// Vertical alignment
				var va = c.GetEffectiveElement<TableCellVerticalAlignment>();
				if (va?.Val != null)
				{
					if (va.Val.Value == TableVerticalAlignmentValues.Top)
						cell.VerticalAlignment = Text.Element.ALIGN_TOP;
					else if (va.Val.Value == TableVerticalAlignmentValues.Bottom)
						cell.VerticalAlignment = Text.Element.ALIGN_BOTTOM;
					else if (va.Val.Value == TableVerticalAlignmentValues.Center)
						cell.VerticalAlignment = Text.Element.ALIGN_MIDDLE;
					else
						cell.VerticalAlignment = cell.VerticalAlignment;
				}

				TableBuilder.BuildTableCellBorder(cellHelper, cell, topPadding, bottomPadding);
			}

			// process Word elements
			var elements = new List<Text.IElement>();
			if (cellHelper.Cell != null)
			{
				_ = tableHelper.GetCellWidth(cellHelper) - cell.PaddingLeft - cell.PaddingRight;
				var list = cellHelper.Cell.Elements().ToList();
				Text.Paragraph? previousParagraph = null;
				for (var i = 0; i < list.Count; i++)
				{
					var e = Dispatcher(list, i, previousParagraph);
					previousParagraph = e as Text.Paragraph;
					if (e != null)
						elements.Add(e);
				}
			}

			TableBuilder.BuildCellHeight(cell, elements, minRowHeight, exactRowHeight);

			// TODO: improve this
			cell.PaddingLeft += 0.5f;
			cell.PaddingTop += 1.0f;

			pt.AddCell(cell);
		}

		return TableBuilder.BuildIndentation(table, pt);
	}

	/// <summary>
	/// Builds the default header from the given section properties
	/// </summary>
	/// <param name="sectionProperties">The section properties</param>
	/// <returns>List of elements</returns>
	List<Text.IElement> BuildHeader(OpenXmlElement sectionProperties)
	{
		var ret = new List<Text.IElement>();

		// Get id
		var reference = sectionProperties.Descendants<HeaderReference>()
			.FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.Default);
		if (reference == null) return ret;
		var id = reference.Id?.Value ?? "";

		// Find element by id
		var headerFooter =
			docxDocument.MainDocumentPart?.HeaderParts.FirstOrDefault(c =>
				docxDocument.MainDocumentPart.GetIdOfPart(c) == id);
		if (headerFooter?.Header == null) return ret;
		var list = headerFooter.Header.Elements().ToList();
		Text.Paragraph? pgPrevious = null;
		for (var i = 0; i < list.Count; i++)
		{
			var e = Dispatcher(list, i, pgPrevious);
			pgPrevious = e as Text.Paragraph;
			if (e == null) continue;

			ret.Add(e);
		}

		return ret;
	}

	/// <summary>
	/// Builds the default footer from the given section properties
	/// </summary>
	/// <param name="sectionProperties">The section properties</param>
	/// <returns>List of elements</returns>
	List<Text.IElement> BuildFooter(OpenXmlElement sectionProperties)
	{
		var ret = new List<Text.IElement>();

		// Get id
		var reference = sectionProperties.Descendants<FooterReference>()
			.FirstOrDefault(x => x.Type?.Value == HeaderFooterValues.Default);
		if (reference == null) return ret;
		var id = reference.Id?.Value ?? "";

		// Find element by id
		var headerFooter =
			docxDocument.MainDocumentPart?.FooterParts.FirstOrDefault(c =>
				docxDocument.MainDocumentPart.GetIdOfPart(c) == id);
		if (headerFooter?.Footer == null) return ret;
		var list = headerFooter.Footer.Elements().ToList();
		Text.Paragraph? pgPrevious = null;
		for (var i = 0; i < list.Count; i++)
		{
			var t = Dispatcher(list, i, pgPrevious);
			pgPrevious = t as Text.Paragraph;
			if (t == null) continue;
			ret.Add(t);
		}

		return ret;
	}
}
