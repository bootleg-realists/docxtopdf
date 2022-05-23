using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;
using BootlegRealists.Reporting.Extension;

namespace BootlegRealists.Reporting.Function;

/// <summary>
/// This class contains OpenXML element functions.
/// </summary>
public static class OpenXmlElementFunction
{
	/// <summary>
	/// Gets the effects for a composite element
	/// </summary>
	/// <param name="compositeElement">Composite element to get the effects for</param>
	/// <param name="bold">[out] Bold effect</param>
	/// <param name="italic">[out] Italic effect</param>
	/// <param name="strike">[out] Strike effect</param>
	/// <param name="caps">[out] Caps effect</param>
	/// <param name="underline">[out] Underline effect</param>
	/// <param name="verticalAlignment">[out] Superscript/subscript effect</param>
	/// <param name="fontSize">[out] Font size (in points)</param>
	/// <param name="fontSizeComplexScript">[out] Complex script font size (in points)</param>
	/// <param name="color">[out] Color</param>
	public static void GetCompositeElementEffects(OpenXmlElement compositeElement, out Bold? bold, out Italic? italic,
		out Strike? strike, out Caps? caps, out Underline? underline, out VerticalAlignment verticalAlignment, out float fontSize,
		out float fontSizeComplexScript, out Color? color)
	{
		const float defaultFontSize = 11.0f;
		bold = compositeElement.GetEffectiveElement<Bold>();
		italic = compositeElement.GetEffectiveElement<Italic>();
		strike = compositeElement.GetEffectiveElement<Strike>();
		caps = compositeElement.GetEffectiveElement<Caps>();
		underline = compositeElement.GetEffectiveElement<Underline>();
		verticalAlignment = VerticalAlignment.None;
		var verticalTextAlignment = compositeElement.GetEffectiveElement<VerticalTextAlignment>();
		if (verticalTextAlignment != null)
		{
			if (verticalTextAlignment.Val?.Value == VerticalPositionValues.Superscript)
				verticalAlignment = VerticalAlignment.Superscript;
			else if (verticalTextAlignment.Val?.Value == VerticalPositionValues.Subscript)
				verticalAlignment = VerticalAlignment.Subscript;
		}
		var size = compositeElement.GetEffectiveElement<FontSize>();
		fontSize = size?.Val != null ? Converter.HalfPointToPoint(size.Val.Value) : defaultFontSize;
		var csSize = compositeElement.GetEffectiveElement<FontSizeComplexScript>();
		fontSizeComplexScript =
			csSize?.Val != null ? Converter.HalfPointToPoint(csSize.Val.Value) : defaultFontSize;
		color = compositeElement.GetEffectiveElement<Color>();
	}
}
