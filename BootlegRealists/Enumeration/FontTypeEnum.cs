namespace BootlegRealists.Reporting.Enumeration;

/// <summary>
/// Enumeration of font types.
/// </summary>
public enum FontTypeEnum
{
	// Complex Script introduction:
	//   https://xmlgraphics.apache.org/fop/1.1/complexscripts.html
	//   http://jrgraphix.net/research/unicode_blocks.php

	Ascii,
	ComplexScript, // complex script, e.g. Thai, Arabic, Hebrew, and right-to-left languages
	EastAsian,
	HighAnsi,
	Unknown
}
