using BootlegRealists.Reporting.Enumeration;

namespace BootlegRealists.Reporting;

/// <summary>
/// Font type information, includes the name of Unicode block, and the font type should be used for this Unicode block.
/// </summary>
public class FontTypeInfo
{
	/// <summary>
	/// The font type (e.g. ComplexScript or EastAsian) of the Unicode block.
	/// </summary>
	public FontTypeEnum FontType { get; set; } = FontTypeEnum.Unknown;

	/// <summary>
	/// Name of the Unicode block.
	/// </summary>
	public string Name { get; set; } = "";

	/// <summary>
	/// If the value of the hint attribute is eastAsia then East Asian font is used, otherwise High ANSI font is used.
	/// </summary>
	public bool UseEastAsiaIfhintIsEastAsia { get; set; }
}

internal class UnicodeBlockRange
{
	public readonly int Begin;
	public readonly int End;

	public UnicodeBlockRange(int begin, int end)
	{
		this.Begin = begin;
		this.End = end;
	}
}

/// <summary>
/// For internal usage, store the Unicode block related information.
/// </summary>
internal class UnicodeBlock : FontTypeInfo
{
	readonly List<UnicodeBlockRange> ranges = new();

	public UnicodeBlock(string name, IEnumerable<UnicodeBlockRange> blocks, FontTypeEnum fontType, bool useEastAsiaIfhintIsEastAsia)
	{
		Name = name;
		ranges.AddRange(blocks);
		FontType = fontType;
		UseEastAsiaIfhintIsEastAsia = useEastAsiaIfhintIsEastAsia;
	}

	/// <summary>
	/// Check the target Unicode value belongs to this code point or not.
	/// </summary>
	/// <param name="unicode">Target Unicode value.</param>
	/// <returns>Return true means the target Unicode value belongs to this code point, otherwise return false.</returns>
	public bool IsIn(int unicode)
	{
		return ranges.Find(r => unicode >= r.Begin && unicode <= r.End) != null;
	}
}

/// <summary>
/// This class is used to detect code points.
/// </summary>
public static class CodePointRecognizer
{
	// https://social.msdn.microsoft.com/Forums/en-US/1bf1f185-ee49-4314-94e7-f4e1563b5c00/finding-which-font-is-to-be-used-to-displaying-a-character-from-pptx-xml?forum=os_binaryfile
	// Unicode character in a run, the font slot can be determined using the following two-step methodology:
	//   1. Use the table below to decide the classification of the content, based on its Unicode code point.
	//   2. If, after the first step, the character falls into East Asian classification and the value of the 
	//      hint attribute is eastAsia, then the character should use East Asian font slot
	//      1. Otherwise, if there is <w:cs/> or <w:rtl/> in this run, then the character should use Complex 
	//         Script font slot, regardless of its Unicode code point.
	//         1. Otherwise, the character is decided using the font slot that is corresponding to the 
	//            classification in the table above.
	// Once the font slot for the run has been determined using the above steps, the appropriate formatting 
	// elements (either complex script or non-complex script) will affect the content.

	static readonly List<UnicodeBlock> Blocks = new(new[]
	{
		new UnicodeBlock("Basic Latin",
			new[] {new UnicodeBlockRange(0x0000, 0x007F)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Latin-1 Supplement",
			new[] {new UnicodeBlockRange(0x00A0, 0x00FF)}, FontTypeEnum.HighAnsi,
			false), // TODO: exception not implemented
		new UnicodeBlock("Latin Extended-A",
			new[] {new UnicodeBlockRange(0x0100, 0x017F)}, FontTypeEnum.HighAnsi,
			false), // TODO: exception not implemented
		new UnicodeBlock("Latin Extended-B",
			new[] {new UnicodeBlockRange(0x0180, 0x024F)}, FontTypeEnum.HighAnsi,
			false), // TODO: exception not implemented
		new UnicodeBlock("IPA Extensions",
			new[] {new UnicodeBlockRange(0x0250, 0x02AF)}, FontTypeEnum.HighAnsi,
			false), // TODO: exception not implemented
		new UnicodeBlock("Spacing Modifier Letters",
			new[] {new UnicodeBlockRange(0x02B0, 0x02FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Combining Diacritical Marks",
			new[] {new UnicodeBlockRange(0x0300, 0x036F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Greek",
			new[] {new UnicodeBlockRange(0x0370, 0x03CF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Cyrillic",
			new[] {new UnicodeBlockRange(0x0400, 0x04FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Hebrew",
			new[] {new UnicodeBlockRange(0x0590, 0x05FF)}, FontTypeEnum.Ascii, false),

		new UnicodeBlock("Arabic",
			new[] {new UnicodeBlockRange(0x0600, 0x06FF)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Syriac",
			new[] {new UnicodeBlockRange(0x0700, 0x074F)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Arabic Supplement",
			new[] {new UnicodeBlockRange(0x0750, 0x077F)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Thaana",
			new[] {new UnicodeBlockRange(0x0780, 0x07BF)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Hangul Jamo",
			new[] {new UnicodeBlockRange(0x1100, 0x11FF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Latin Extended Additional",
			new[] {new UnicodeBlockRange(0x1E00, 0x1EFF)}, FontTypeEnum.HighAnsi,
			false), // TODO: exception not implemented
		new UnicodeBlock("Greek Extended",
			new[] {new UnicodeBlockRange(0x1F00, 0x1FFF)}, FontTypeEnum.HighAnsi, false),
		new UnicodeBlock("General Punctuation",
			new[] {new UnicodeBlockRange(0x2000, 0x206F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Superscripts and Subscripts",
			new[] {new UnicodeBlockRange(0x2070, 0x209F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Currency Symbols",
			new[] {new UnicodeBlockRange(0x20A0, 0x20CF)}, FontTypeEnum.HighAnsi, true),

		new UnicodeBlock("Combining Diacritical Marks for Symbols",
			new[] {new UnicodeBlockRange(0x20D0, 0x20FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Letter-like Symbols",
			new[] {new UnicodeBlockRange(0x2100, 0x214F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Number Forms",
			new[] {new UnicodeBlockRange(0x2150, 0x218F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Arrows",
			new[] {new UnicodeBlockRange(0x2190, 0x21FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Mathematical Operators",
			new[] {new UnicodeBlockRange(0x2200, 0x22FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Miscellaneous Technical",
			new[] {new UnicodeBlockRange(0x2300, 0x23FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Control Pictures",
			new[] {new UnicodeBlockRange(0x2400, 0x243F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Optical Character Recognition",
			new[] {new UnicodeBlockRange(0x2440, 0x245F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Enclosed Alphanumerics",
			new[] {new UnicodeBlockRange(0x2460, 0x24FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Box Drawing",
			new[] {new UnicodeBlockRange(0x2500, 0x257F)}, FontTypeEnum.HighAnsi, true),

		new UnicodeBlock("Block Elements",
			new[] {new UnicodeBlockRange(0x2580, 0x259F)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Geometric Shapes",
			new[] {new UnicodeBlockRange(0x25A0, 0x25FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Miscellaneous Symbols",
			new[] {new UnicodeBlockRange(0x2600, 0x26FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Dingbats",
			new[] {new UnicodeBlockRange(0x2700, 0x27BF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("CJK Radicals Supplement",
			new[] {new UnicodeBlockRange(0x2E80, 0x2EFF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Kangxi Radicals",
			new[] {new UnicodeBlockRange(0x2F00, 0x2FDF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Ideographic Description Characters",
			new[] {new UnicodeBlockRange(0x2FF0, 0x2FFF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("CJK Symbols and Punctuation",
			new[] {new UnicodeBlockRange(0x3000, 0x303F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Hiragana",
			new[] {new UnicodeBlockRange(0x3040, 0x309F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Katakana",
			new[] {new UnicodeBlockRange(0x30A0, 0x30FF)}, FontTypeEnum.EastAsian, false),

		new UnicodeBlock("Bopomofo",
			new[] {new UnicodeBlockRange(0x3100, 0x312F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Hangul Compatibility Jamo",
			new[] {new UnicodeBlockRange(0x3130, 0x318F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Kanbun",
			new[] {new UnicodeBlockRange(0x3190, 0x319F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Enclosed CJK Letters and Months",
			new[] {new UnicodeBlockRange(0x3200, 0x32FF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("CJK Compatibility",
			new[] {new UnicodeBlockRange(0x3300, 0x33FF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("CJK Unified Ideographs Extension A",
			new[] {new UnicodeBlockRange(0x3400, 0x4DBF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("CJK Unified Ideographs",
			new[] {new UnicodeBlockRange(0x4E00, 0x9FAF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Yi Syllables",
			new[] {new UnicodeBlockRange(0xA000, 0xA48F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Yi Radicals",
			new[] {new UnicodeBlockRange(0xA490, 0xA4CF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Hangul Syllables",
			new[] {new UnicodeBlockRange(0xAC00, 0xD7AF)}, FontTypeEnum.EastAsian, false),

		new UnicodeBlock("High Surrogates",
			new[] {new UnicodeBlockRange(0xD800, 0xDB7F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("High Use Surrogates",
			new[] {new UnicodeBlockRange(0xDB80, 0xDBFF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Low Surrogates",
			new[] {new UnicodeBlockRange(0xDC00, 0xDFFF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Use Area",
			new[] {new UnicodeBlockRange(0xE000, 0xF8FF)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("CJK Compatibility Ideographs",
			new[] {new UnicodeBlockRange(0xF900, 0xFAFF)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Alphabetic Presentation Forms1",
			new[] {new UnicodeBlockRange(0xFB00, 0xFB1C)}, FontTypeEnum.HighAnsi, true),
		new UnicodeBlock("Alphabetic Presentation Forms2",
			new[] {new UnicodeBlockRange(0xFB1D, 0xFB4F)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Arabic Presentation Forms-A",
			new[] {new UnicodeBlockRange(0xFB50, 0xFDFF)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("CJK Compatibility Forms",
			new[] {new UnicodeBlockRange(0xFE30, 0xFE4F)}, FontTypeEnum.EastAsian, false),
		new UnicodeBlock("Small Form Variants",
			new[] {new UnicodeBlockRange(0xFE50, 0xFE6F)}, FontTypeEnum.EastAsian, false),

		new UnicodeBlock("Arabic Presentation Forms-B",
			new[] {new UnicodeBlockRange(0xFE70, 0xFEFE)}, FontTypeEnum.Ascii, false),
		new UnicodeBlock("Halfwidth and Fullwidth Forms",
			new[] {new UnicodeBlockRange(0xFF00, 0xFFEF)}, FontTypeEnum.EastAsian, false)
	});

	/// <summary>
	/// Get font type information by a Unicode value. Font type information indicates the character should be display in
	/// ASCII/Complex Script/EastAsian/HighANSI.
	/// </summary>
	/// <param name="unicode">Unicode value of the character.</param>
	/// <returns>Return font type information.</returns>
	public static FontTypeInfo GetFontType(int unicode)
	{
		var ret = new FontTypeInfo();

		var unicodeBlock = Blocks.Find(block => block.IsIn(unicode));
		if (unicodeBlock == null) return ret;
		ret.Name = unicodeBlock.Name;
		ret.FontType = unicodeBlock.FontType;
		ret.UseEastAsiaIfhintIsEastAsia = unicodeBlock.UseEastAsiaIfhintIsEastAsia;

		return ret;
	}
}
