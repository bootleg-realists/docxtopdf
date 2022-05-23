
namespace BootlegRealists.Reporting.Enumeration;

/// <summary>
/// This enumeration contains font styles
/// </summary>
[Flags]
internal enum FontStyle : ushort
{
	Unknown = 0,

	Italic = 1,
	Bold = 1 << 1,
	Regular = 1 << 2,
	Oblique = 1 << 3
}
