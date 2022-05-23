using System.Globalization;

namespace BootlegRealists.Reporting;

internal static class Tools
{
	/// <summary>
	/// Roman numerals
	/// </summary>
	static readonly string[][] RomanNumerals =
	{
		new[] {"", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"}, // ones
		new[] {"", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"}, // tens
		new[] {"", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM"}, // hundreds
		new[] {"", "M", "MM", "MMM"} // thousands
	};

	/// <summary>
	/// Taiwanese numerals
	/// </summary>
	static readonly string[][] TaiwaneseNumerals =
	{
		new[] {"", "一", "二", "三", "四", "五", "六", "七", "八", "九"}, // ones
		new[] {"零", "一十", "二十", "三十", "四十", "五十", "六十", "七十", "八十", "九十"}, // tens
		new[] {"零", "一百", "二百", "三百", "四百", "五百", "六百", "七百", "八百", "九百"}, // hundreds
		new[] {"零", "一千", "二千", "三千", "四千", "五千", "六千", "七千", "八千", "九千"} // thousands
	};

	/// <summary>
	/// Get color's brightness.
	/// </summary>
	/// <param name="color">Color string, e.g. FF0000.</param>
	/// <returns></returns>
	public static float RgbBrightness(string color)
	{
		if (string.IsNullOrEmpty(color) || color == "auto") // TODO: handle auto color
			color = "0";

		var rgb = Convert.ToInt32(color, 16);
		return RgbBrightness((rgb & 0xff0000) >> 16, (rgb & 0xff00) >> 8, rgb & 0xff);
	}

	/// <summary>
	/// Get color's brightness.
	/// </summary>
	/// <param name="r">R</param>
	/// <param name="g">G</param>
	/// <param name="b">B</param>
	/// <returns></returns>
	static float RgbBrightness(int r, int g, int b)
	{
		if (r > 255)
			r = 255;
		if (r < 0)
			r = 0;

		if (g > 255)
			g = 255;
		if (g < 0)
			g = 0;

		if (b > 255)
			b = 255;
		if (b < 0)
			b = 0;

		// http://stackoverflow.com/questions/596216/formula-to-determine-brightness-of-rgb-color
		// http://www.codeproject.com/Articles/19045/Manipulating-colors-in-NET-Part-1
		//return (float)(0.2126f * r + 0.7152 * g + 0.0722 * b);

		// http://www.w3.org/TR/AERT#color-contrast
		return (float)(0.299f * r + 0.587 * g + 0.114 * b);
	}

	/// <summary>
	/// Convert percentage string (e.g. "50%" or "2500") to the value between 0~1.
	/// </summary>
	/// <param name="str">Percentage string, e.g. "50%" or "2500".</param>
	/// <returns>Return a float value between 0~1.</returns>
	public static float Percentage(string str)
	{
		if (string.IsNullOrEmpty(str)) return 0.0f;
		var s = str.Trim();
		if (s.EndsWith("%", StringComparison.Ordinal))
		{
			if (float.TryParse(s[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
				return v / 100.0f;
		}
		else
		{
			if (float.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
				return v / 5000.0f;
		}

		return 0.0f;
	}

	static string IntToAnything(IReadOnlyList<string[]> numerals, int number, bool ignoreZeroes)
	{
		// split integer string into array and reverse array
		var intArr = new string(number.ToString(CultureInfo.InvariantCulture).Reverse().ToArray());
		var ret = "";
		var end = 0;

		if (ignoreZeroes)
		{
			while (intArr[end] == '0' && end < intArr.Length - 1)
				end++;
		}

		// starting with the highest place (for 3046, it would be the thousands
		// place, or 3), get the roman numeral representation for that place
		// and add it to the final roman numeral string
		for (var i = intArr.Length - 1; i >= end; i--) ret += numerals[i][intArr[i] - '0'];

		return ret;
	}

	/// <summary>
	/// Convert integer to Roman numeral expression.
	/// </summary>
	/// <param name="number">Integer.</param>
	/// <param name="uppercase">True for upper case, false for lower case.</param>
	/// <returns></returns>
	public static string IntToRoman(int number, bool uppercase)
	{
		return uppercase
			? IntToAnything(RomanNumerals, number, false)
			: IntToAnything(RomanNumerals, number, false).ToLower(CultureInfo.InvariantCulture);
	}

	/// <summary>
	/// Convert integer to Taiwanese numeral expression.
	/// </summary>
	/// <param name="number">Integer.</param>
	/// <returns></returns>
	public static string IntToTaiwanese(int number)
	{
		return IntToAnything(TaiwaneseNumerals, number, true);
	}
}
