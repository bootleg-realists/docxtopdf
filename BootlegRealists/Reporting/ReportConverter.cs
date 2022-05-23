using BootlegRealists.Reporting.Interface;

namespace BootlegRealists.Reporting;

/// <summary>
/// Base class for the report converter.
/// </summary>
public abstract class ReportConverter : IReportConverter
{
	/// <inheritdoc />
	public abstract void Execute(Stream inputStream, Stream outputStream,
		IDictionary<string, string>? runProperties = null);
}
