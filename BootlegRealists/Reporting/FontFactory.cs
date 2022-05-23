using System.Drawing.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using Typography.OpenFont;
using Typography.OpenFont.Extensions;
using FontFamily = System.Drawing.FontFamily;
using Pdf = iTextSharp.text.pdf;
using Text = iTextSharp.text;

// TODO: handle sectPr w:docGrid w:linePitch (how many lines per page)
// TODO: handle tblLook $17.7.6 (conditional formatting), it will define first row/first column/...etc styles in styles.xml

namespace BootlegRealists.Reporting;

/// <summary>
/// This class manages the fonts in the system.
/// </summary>
public static class FontFactory
{
	/// <summary>
	/// Paths of the fonts
	/// </summary>
	static readonly List<string> FontFolderPaths = GetFontFolderPaths();

	static readonly List<(string Path, string Name, FontStyle FontStyle)> Fonts = new();

	static readonly string[] IllegalFontNames =
	{
		"eastAsia", "cs", "hAnsi", "ascii", "majorBidi", "minorBidi", "majorHAnsi", "minorHAnsi",
		"majorEastAsia", "minorEastAsia"
	};

	/// <summary>
	/// The constructor of FontFactory.
	/// </summary>
	static FontFactory()
	{
		var files = new List<string>();
		foreach (var fontFolderPath in FontFolderPaths)
		{
			files.AddRange(Directory.GetFiles(fontFolderPath, "*.*", SearchOption.AllDirectories).Where(x =>
			{
				var ext = Path.GetExtension(x).ToLower(CultureInfo.InvariantCulture);
				return ext switch
				{
					".ttc" or ".otc" or ".ttf" or ".otf" or ".woff" or ".woff2" => true,
					_ => false
				};
			}).ToList());
		}
		files.Sort();
		foreach (var file in files)
		{
			var (fontName, fontStyle) = GetFontProperties(file);
			Fonts.Add((file + (file.EndsWith(".ttc", StringComparison.Ordinal) ? ",0" : ""), fontName, fontStyle));
		}
	}

	/// <summary>
	/// Gets the list of font folder paths
	/// </summary>
	/// <returns>The list</returns>
	static List<string> GetFontFolderPaths()
	{
		var result = new List<string>();
		if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
		{
			result.Add("/System/Library/Fonts");
			result.Add("/Library/Fonts");
		}

		result.Add(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
		return result;
	}

	/// <summary>
	/// Get font properties from file, e.g. "Arial", Regular from @"C:\windows\Fonts\arial.ttf".
	/// </summary>
	/// <param name="fileName">The full path of the font file.</param>
	/// <returns>The font properties in English-United States.</returns>
	static (string, FontStyle) GetFontProperties(string fileName)
	{
		var fontStyle = FontStyle.Unknown;
		using var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
		try
		{
			var reader = new OpenFontReader();
			var typeface = reader.Read(fs);
			fontStyle = (FontStyle)typeface.TranslateOS2FontStyle();
		}
		catch (NullReferenceException)
		{
		}
		catch (NotSupportedException)
		{
		}
		catch (NotImplementedException)
		{
		}

		var collection = new PrivateFontCollection();
		collection.AddFontFile(fileName);
		return (collection.Families.Length > 0 ? collection.Families[0].GetName(0x0409) : "", fontStyle);
	}

	/// <summary>
	/// Search for font file by font properties.
	/// </summary>
	/// <param name="name">Font name (e.g. Arial or 標楷體).</param>
	/// <param name="bold">True if font is bold</param>
	/// <param name="italic">True if font is italic</param>
	/// <returns>The full path of font file or null if can't find and the exact match (or not).</returns>
	static (string Path, bool ExactMatch) SearchByProperties(string name, bool bold, bool italic)
	{
		if (string.IsNullOrEmpty(name) || IllegalFontNames.Contains(name))
			return ("", false);

		try
		{
			var fontFamily = new FontFamily(name);
			var enName = fontFamily.GetName(0x0409);

			// Exact math
			var fontInfo = Fonts.Find(x =>
				string.Equals(x.Name, enName, StringComparison.OrdinalIgnoreCase) &&
				(x.FontStyle & FontStyle.Bold) != 0 == bold && (x.FontStyle & FontStyle.Italic) != 0 == italic);
			if (fontInfo != default)
				return (fontInfo.Path, true);
			// First match
			fontInfo = Fonts.Find(x => string.Equals(x.Name, enName, StringComparison.OrdinalIgnoreCase));
			if (fontInfo != default)
				return (fontInfo.Path, false);
		}
		catch (ArgumentException)
		{
		}

		return ("", false);
	}

	/// <summary>
	/// Create iTextSharp.text.pdf.BaseFont by font name. This method doesn't configure font size, color or any other
	/// properties for BaseFont.
	/// </summary>
	/// <param name="fontName">Font name (e.g. Arial or 標楷體)</param>
	/// <param name="bold">True if font is bold</param>
	/// <param name="italic">True if font is italic</param>
	/// <returns>Return BaseFontEx object or null if can't find font.</returns>
	public static BaseFontEx? CreateBaseFont(string fontName, bool bold, bool italic)
	{
		if (string.IsNullOrEmpty(fontName))
			return null;

		const string encoding = "";
		const bool embedded = Pdf.BaseFont.EMBEDDED;

		// Special case: symbol font
		if (fontName == "Symbol")
		{
			var result = BaseFontEx.CreateFont(fontName, encoding, embedded);
			result.ExactMatch = true;
			return result;
		}

		var (fontFilePath, exactMatch) = SearchByProperties(fontName, bold, italic);
		if (string.IsNullOrEmpty(fontFilePath))
			return null;
		var font = BaseFontEx.CreateFont(fontFilePath, encoding, embedded);
		font.ExactMatch = exactMatch;
		return font;
	}

	/// <summary>
	/// Create iTextSharp.text.Font by specifying the BaseFont, font size, and other font properties.
	/// </summary>
	/// <param name="baseFont">Font to use as base</param>
	/// <param name="fontSize">Font size in points.</param>
	/// <param name="bold">Bold property.</param>
	/// <param name="italic">Italic property.</param>
	/// <param name="strike">Strike property.</param>
	/// <param name="color">Color.</param>
	/// <returns>The font</returns>
	public static Text.Font CreateFont(BaseFontEx baseFont, float fontSize, Bold? bold, Italic? italic, Strike? strike, Color? color)
	{
		var result = new Text.Font(baseFont);

		var rgb = 0;
		if (color?.Val != null && color.Val.Value != "auto")
			rgb = Convert.ToInt32(color.Val.Value, 16);

		result.SetStyle((!baseFont.ExactMatch && Converter.OnOffToBool(bold) ? Text.Font.BOLD : 0) |
		                (!baseFont.ExactMatch && Converter.OnOffToBool(italic) ? Text.Font.ITALIC : 0) |
		                (Converter.OnOffToBool(strike) ? Text.Font.STRIKETHRU : 0));
		if (rgb != 0)
			result.SetColor((rgb & 0xff0000) >> 16, (rgb & 0xff00) >> 8, rgb & 0xff);

		if (fontSize > 0f)
			result.Size = fontSize;

		return result;
	}
}
