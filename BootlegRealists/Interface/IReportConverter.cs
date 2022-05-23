
namespace BootlegRealists.Reporting.Interface;

/// <summary>
/// Interface for the report converter
/// </summary>
public interface IReportConverter
{
	/// <summary>
	/// Converts the input stream (containing the source report) to the output stream (which will contain the destination
	/// report).
	/// </summary>
	/// <param name="inputStream">Input stream (source report)</param>
	/// <param name="outputStream">Output stream (destination report)</param>
	/// <param name="runProperties">Properties to use destination report</param>
	void Execute(Stream inputStream, Stream outputStream, IDictionary<string, string>? runProperties = null);
}
