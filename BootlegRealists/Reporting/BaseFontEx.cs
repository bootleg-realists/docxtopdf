using iTextSharp.text.pdf;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class extends the BaseFont class by introducing new methods and properties
/// </summary>
public sealed class BaseFontEx
{
	/// <summary>
	/// Creates a new font. This font can be one of the 14 built in types, a Type1 font
	/// referred to by an AFM or PFM file, a TrueType font (simple or collection) or
	/// a CJK font from the Adobe Asian Font Pack. TrueType fonts and CJK fonts can have
	/// an optional style modifier appended to the name. These modifiers are: Bold, Italic
	/// and BoldItalic. An example would be "STSong-Light,Bold". Note that this modifiers
	/// do not work if the font is embedded. Fonts in TrueType collections are addressed
	/// by index such as "msgothic.ttc,1". This would get the second font (indexes start
	/// at 0), in this case "MS PGothic". The fonts are cached and if they already exist
	/// they are extracted from the cache, not parsed again. Besides the common encodings
	/// described by name, custom encodings can also be made. These encodings will only
	/// work for the single byte fonts Type1 and TrueType. The encoding string starts
	/// with a '#' followed by "simple" or "full". If "simple" there is a decimal for
	/// the first character position and then a list of hex values representing the Unicode
	/// codes that compose that encoding. The "simple" encoding is recommended for TrueType
	/// fonts as the "full" encoding risks not matching the character with the right
	/// glyph if not done with care. The "full" encoding is specially aimed at Type1
	/// fonts where the glyphs have to be described by non standard names like the Tex
	/// math fonts. Each group of three elements compose a code position: the one byte
	/// code order in decimal or as 'x' (x cannot be the space), the name and the Unicode
	/// character used to access the glyph. The space must be assigned to character position
	/// 32 otherwise text justification will not work. Example for a "simple" encoding
	/// that includes the Unicode character space, A, B and ecyrillic: "# simple 32 0020
	/// 0041 0042 0454" Example for a "full" encoding for a Type1 Tex font: "# full 'A'
	/// nottriangeqlleft 0041 'B' dividemultiply 0042 32 space 0020" This method calls:
	/// createFont(name, encoding, embedded, true, null, null); @throws DocumentException
	/// the font is invalid @throws IOException the font file could not be read
	/// </summary>
	/// <param name="name">the name of the font or its location on file</param>
	/// <param name="encoding">the encoding to be applied to this font</param>
	/// <param name="embedded">true if the font is to be embedded in the PDF</param>
	/// <returns>returns a new font. This font may come from the cache</returns>
	public static BaseFontEx CreateFont(string name, string encoding, bool embedded) => new(BaseFont.CreateFont(name, encoding, embedded));
	/// <summary>
	/// Creates a font based on an existing document font. The created font may
	/// not behave as expected, depending on the encoding or subset.
	/// </summary>
	/// <param name="fontRef">the reference to the document font</param>
	/// <returns>the font</returns>
	public static BaseFontEx CreateFont(PrIndirectReference fontRef) => new(BaseFont.CreateFont(fontRef));
	/// <summary>
	/// /// Creates a new font. This font can be one of the 14 built in types, a Type1 font
	/// referred to by an AFM or PFM file, a TrueType font (simple or collection) or
	/// a CJK font from the Adobe Asian Font Pack. TrueType fonts and CJK fonts can have
	/// an optional style modifier appended to the name. These modifiers are: Bold, Italic
	/// and BoldItalic. An example would be "STSong-Light,Bold". Note that this modifiers
	/// do not work if the font is embedded. Fonts in TrueType collections are addressed
	/// by index such as "msgothic.ttc,1". This would get the second font (indexes start
	/// at 0), in this case "MS PGothic". The fonts may or may not be cached depending
	/// on the flag cached . If the byte arrays are present the font will be read from
	/// them instead of the name. A name is still required to identify the font type.
	/// Besides the common encodings described by name, custom encodings can also be
	/// made. These encodings will only work for the single byte fonts Type1 and TrueType.
	/// The encoding string starts with a '#' followed by "simple" or "full". If "simple"
	/// there is a decimal for the first character position and then a list of hex values
	/// representing the Unicode codes that compose that encoding. The "simple" encoding
	/// is recommended for TrueType fonts as the "full" encoding risks not matching the
	/// character with the right glyph if not done with care. The "full" encoding is
	/// specially aimed at Type1 fonts where the glyphs have to be described by non standard
	/// names like the Tex math fonts. Each group of three elements compose a code position:
	/// the one byte code order in decimal or as 'x' (x cannot be the space), the name
	/// and the Unicode character used to access the glyph. The space must be assigned
	/// to character position 32 otherwise text justification will not work. Example
	/// for a "simple" encoding that includes the Unicode character space, A, B and ecyrillic:
	/// "# simple 32 0020 0041 0042 0454" Example for a "full" encoding for a Type1 Tex
	/// font: "# full 'A' nottriangeqlleft 0041 'B' dividemultiply 0042 32 space 0020"
	/// the cache if new, false if the font is always created new an exception if the
	/// font is not recognized. Note that even if true an exception may be thrown in
	/// some circumstances. This parameter is useful for FontFactory that may have to
	/// check many invalid font names before finding the right one is true, otherwise
	/// it will always be created new @throws DocumentException the font is invalid @throws
	/// IOException the font file could not be read @since 2.1.5
	/// </summary>
	/// <param name="name">the name of the font or its location on file</param>
	/// <param name="encoding">the encoding to be applied to this font</param>
	/// <param name="embedded">true if the font is to be embedded in the PDF</param>
	/// <param name="cached">true if the font comes from the cache or is added to</param>
	/// <param name="ttfAfm">the true type font or the afm in a byte array</param>
	/// <param name="pfb">the pfb in a byte array</param>
	/// <param name="noThrow">if true will not throw an exception if the font is not recognized and will return
	/// null, if false will throw</param>
	/// <param name="forceRead">in some cases (TrueTypeFont, Type1Font), the full font file will be read and
	/// kept in memory if forceRead is true</param>
	/// <returns>returns a new font. This font may come from the cache but only if cached</returns>
	public static BaseFontEx CreateFont(string name, string encoding, bool embedded, bool cached, byte[] ttfAfm, byte[] pfb, bool noThrow, bool forceRead) => new(BaseFont.CreateFont(name, encoding, embedded, cached, ttfAfm, pfb, noThrow, forceRead));
	/// <summary>
	/// Creates a new font. This font can be one of the 14 built in types, a Type1 font
	/// referred to by an AFM or PFM file, a TrueType font (simple or collection) or
	/// a CJK font from the Adobe Asian Font Pack. TrueType fonts and CJK fonts can have
	/// an optional style modifier appended to the name. These modifiers are: Bold, Italic
	/// and BoldItalic. An example would be "STSong-Light,Bold". Note that this modifiers
	/// do not work if the font is embedded. Fonts in TrueType collections are addressed
	/// by index such as "msgothic.ttc,1". This would get the second font (indexes start
	/// at 0), in this case "MS PGothic". The fonts may or may not be cached depending
	/// on the flag cached . If the byte arrays are present the font will be read from
	/// them instead of the name. A name is still required to identify the font type.
	/// Besides the common encodings described by name, custom encodings can also be
	/// made. These encodings will only work for the single byte fonts Type1 and TrueType.
	/// The encoding string starts with a '#' followed by "simple" or "full". If "simple"
	/// there is a decimal for the first character position and then a list of hex values
	/// representing the Unicode codes that compose that encoding. The "simple" encoding
	/// is recommended for TrueType fonts as the "full" encoding risks not matching the
	/// character with the right glyph if not done with care. The "full" encoding is
	/// specially aimed at Type1 fonts where the glyphs have to be described by non standard
	/// names like the Tex math fonts. Each group of three elements compose a code position:
	/// the one byte code order in decimal or as 'x' (x cannot be the space), the name
	/// and the Unicode character used to access the glyph. The space must be assigned
	/// to character position 32 otherwise text justification will not work. Example
	/// for a "simple" encoding that includes the Unicode character space, A, B and ecyrillic:
	/// "# simple 32 0020 0041 0042 0454" Example for a "full" encoding for a Type1 Tex
	/// font: "# full 'A' nottriangeqlleft 0041 'B' dividemultiply 0042 32 space 0020"
	/// the cache if new, false if the font is always created new an exception if the
	/// font is not recognized. Note that even if true an exception may be thrown in
	/// some circumstances. This parameter is useful for FontFactory that may have to
	/// check many invalid font names before finding the right one is true, otherwise
	/// it will always be created new @throws DocumentException the font is invalid @throws
	/// IOException the font file could not be read @since 2.0.3
	/// </summary>
	/// <param name="name">the name of the font or its location on file</param>
	/// <param name="encoding">the encoding to be applied to this font</param>
	/// <param name="embedded">true if the font is to be embedded in the PDF</param>
	/// <param name="cached">true if the font comes from the cache or is added to</param>
	/// <param name="ttfAfm">the true type font or the afm in a byte array</param>
	/// <param name="pfb">the pfb in a byte array</param>
	/// <param name="noThrow">if true will not throw an exception if the font is not recognized and will return
	/// null, if false will throw</param>
	/// <returns>returns a new font. This font may come from the cache but only if cached</returns>
	public static BaseFontEx CreateFont(string name, string encoding, bool embedded, bool cached, byte[] ttfAfm, byte[] pfb, bool noThrow) => new(BaseFont.CreateFont(name, encoding, embedded, cached, ttfAfm, pfb, noThrow));
	/// <summary>
	/// Creates a new font. This font can be one of the 14 built in types, a Type1 font
	/// referred to by an AFM or PFM file, a TrueType font (simple or collection) or
	/// a CJK font from the Adobe Asian Font Pack. TrueType fonts and CJK fonts can have
	/// an optional style modifier appended to the name. These modifiers are: Bold, Italic
	/// and BoldItalic. An example would be "STSong-Light,Bold". Note that this modifiers
	/// do not work if the font is embedded. Fonts in TrueType collections are addressed
	/// by index such as "msgothic.ttc,1". This would get the second font (indexes start
	/// at 0), in this case "MS PGothic". The fonts may or may not be cached depending
	/// on the flag cached . If the byte arrays are present the font will be read from
	/// them instead of the name. A name is still required to identify the font type.
	/// Besides the common encodings described by name, custom encodings can also be
	/// made. These encodings will only work for the single byte fonts Type1 and TrueType.
	/// The encoding string starts with a '#' followed by "simple" or "full". If "simple"
	/// there is a decimal for the first character position and then a list of hex values
	/// representing the Unicode codes that compose that encoding. The "simple" encoding
	/// is recommended for TrueType fonts as the "full" encoding risks not matching the
	/// character with the right glyph if not done with care. The "full" encoding is
	/// specially aimed at Type1 fonts where the glyphs have to be described by non standard
	/// names like the Tex math fonts. Each group of three elements compose a code position:
	/// the one byte code order in decimal or as 'x' (x cannot be the space), the name
	/// and the Unicode character used to access the glyph. The space must be assigned
	/// to character position 32 otherwise text justification will not work. Example
	/// for a "simple" encoding that includes the Unicode character space, A, B and ecyrillic:
	/// "# simple 32 0020 0041 0042 0454" Example for a "full" encoding for a Type1 Tex
	/// font: "# full 'A' nottriangeqlleft 0041 'B' dividemultiply 0042 32 space 0020"
	/// the cache if new, false if the font is always created new is true, otherwise
	/// it will always be created new @throws DocumentException the font is invalid @throws
	/// IOException the font file could not be read @since iText 0.80
	/// </summary>
	/// <param name="name">the name of the font or its location on file</param>
	/// <param name="encoding">the encoding to be applied to this font</param>
	/// <param name="embedded">true if the font is to be embedded in the PDF</param>
	/// <param name="cached">true if the font comes from the cache or is added to</param>
	/// <param name="ttfAfm">the true type font or the afm in a byte array</param>
	/// <param name="pfb">the pfb in a byte array</param>
	/// <returns>returns a new font. This font may come from the cache but only if cached</returns>
	public static BaseFontEx CreateFont(string name, string encoding, bool embedded, bool cached, byte[] ttfAfm, byte[] pfb) => new(BaseFont.CreateFont(name, encoding, embedded, cached, ttfAfm, pfb));
	/// <summary>
	/// Creates a new font. This font can be one of the 14 built in types, a Type1 font
	/// referred to by an AFM or PFM file, a TrueType font (simple or collection) or
	/// a CJK font from the Adobe Asian Font Pack. TrueType fonts and CJK fonts can have
	/// an optional style modifier appended to the name. These modifiers are: Bold, Italic
	/// and BoldItalic. An example would be "STSong-Light,Bold". Note that this modifiers
	/// do not work if the font is embedded. Fonts in TrueType collections are addressed
	/// by index such as "msgothic.ttc,1". This would get the second font (indexes start
	/// at 0), in this case "MS PGothic". The fonts are cached and if they already exist
	/// they are extracted from the cache, not parsed again. Besides the common encodings
	/// described by name, custom encodings can also be made. These encodings will only
	/// work for the single byte fonts Type1 and TrueType. The encoding string starts
	/// with a '#' followed by "simple" or "full". If "simple" there is a decimal for
	/// the first character position and then a list of hex values representing the Unicode
	/// codes that compose that encoding. The "simple" encoding is recommended for TrueType
	/// fonts as the "full" encoding risks not matching the character with the right
	/// glyph if not done with care. The "full" encoding is specially aimed at Type1
	/// fonts where the glyphs have to be described by non standard names like the Tex
	/// math fonts. Each group of three elements compose a code position: the one byte
	/// code order in decimal or as 'x' (x cannot be the space), the name and the Unicode
	/// character used to access the glyph. The space must be assigned to character position
	/// 32 otherwise text justification will not work. Example for a "simple" encoding
	/// that includes the Unicode character space, A, B and ecyrillic: "# simple 32 0020
	/// 0041 0042 0454" Example for a "full" encoding for a Type1 Tex font: "# full 'A'
	/// nottriangeqlleft 0041 'B' dividemultiply 0042 32 space 0020" This method calls:
	/// createFont(name, encoding, embedded, true, null, null); @throws DocumentException
	/// the font is invalid @throws IOException the font file could not be read @since
	/// 2.1.5
	/// </summary>
	/// <param name="name">the name of the font or its location on file</param>
	/// <param name="encoding">the encoding to be applied to this font</param>
	/// <param name="embedded">true if the font is to be embedded in the PDF</param>
	/// <param name="forceRead">in some cases (TrueTypeFont, Type1Font), the full font file will be read and
	/// kept in memory if forceRead is true</param>
	/// <returns>returns a new font. This font may come from the cache</returns>
	public static BaseFontEx CreateFont(string name, string encoding, bool embedded, bool forceRead) => new(BaseFont.CreateFont(name, encoding, embedded, forceRead));
	/// <summary>
	/// Creates a new font. This will always be the default Helvetica font (not embedded).
	/// This method is introduced because Helvetica is used in many examples. @throws
	/// IOException This shouldn't occur ever @throws DocumentException This shouldn't
	/// occur ever @since 2.1.1
	/// </summary>
	/// <returns>a BaseFont object (Helvetica, Winansi, not embedded)</returns>
	public static BaseFontEx CreateFont() => new(BaseFont.CreateFont());
	/// <summary>
	/// Converts an instance of the <see cref="BaseFontEx"/> class to a <see cref="iTextSharp.text.pdf.BaseFont"/>
	/// </summary>
	/// <param name="obj">Instance to convert</param>
	/// <returns>The base font</returns>
	public static implicit operator BaseFont(BaseFontEx obj) => obj.font;
	/// <summary>
	/// Gets/sets the flag for if the font is an exact match or not.
	/// </summary>
	public bool ExactMatch { get; set; }
	/// <summary>
	/// Initializes a new instance of the <see cref="BaseFontEx"/> class.
	/// </summary>
	/// <param name="font">Font to set</param>
	BaseFontEx(BaseFont font) => this.font = font;
	/// <summary>
	/// Gets/sets the base font
	/// </summary>
	readonly BaseFont font;
}
