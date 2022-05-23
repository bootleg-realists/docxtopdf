using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting;

/// <summary>
/// This class contains converter functions.
/// </summary>
public static class Converter
{
	/// <summary>
	/// Converts a hundredth of a character to points
	/// </summary>
	/// <param name="f">Hundredth of a character to convert</param>
	/// <param name="fontSizePoint">Size of the reference font in points</param>
	/// <returns>Points</returns>
	public static float HundredthOfCharacterToPoint(float f, float fontSizePoint)
	{
		return f / 100.0f * fontSizePoint;
	}

	/// <summary>
	/// Converts a one eighth of a point to points
	/// </summary>
	/// <param name="f">One eighth of a point to convert</param>
	/// <returns>Points</returns>
	public static float OneEighthPointToPoint(float f)
	{
		return f / 8.0f;
	}

	/// <summary>
	/// Converts a half-point to points
	/// </summary>
	/// <param name="s">String containing half-point to convert</param>
	/// <returns>Points</returns>
	public static float HalfPointToPoint(string? s)
	{
		return float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
			? HalfPointToPoint(v)
			: float.NaN;
	}

	/// <summary>
	/// Converts a twip to points
	/// </summary>
	/// <param name="f">Twip to convert</param>
	/// <returns>Points</returns>
	public static float TwipToPoint(float f)
	{
		return f / 20.0f;
	}

	/// <summary>
	/// Converts a twip to points
	/// </summary>
	/// <param name="s">String containing twip to convert</param>
	/// <returns>Points</returns>
	public static float TwipToPoint(string? s)
	{
		return float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
			? TwipToPoint(v)
			: float.NaN;
	}

	/// <summary>
	/// Converts an on/off object to bool
	/// </summary>
	/// <param name="v">Object to convert</param>
	/// <param name="defaultValue">Default value to use if the internal value is null</param>
	/// <returns>The bool result</returns>
	public static bool OnOffToBool(OnOffType? v, bool defaultValue = true)
	{
		if (v == null) return false;
		return v.Val == (object?)null ? defaultValue : v.Val.Value;
	}
	/// <summary>
	/// Converts a half-point to points
	/// </summary>
	/// <param name="f">Half-point to convert</param>
	/// <returns>Points</returns>
	static float HalfPointToPoint(float f)
	{
		return f / 2.0f;
	}
}
