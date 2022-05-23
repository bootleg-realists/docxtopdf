namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains IList extension methods.
/// </summary>
public static class IListExtension
{
    /// <summary>Adds an object to the end of the list (if it is not null).</summary>
    /// <param name="obj">The instance to act on.</param>
    /// <param name="item">The object to be added to the end of the list.</param>
    public static void AddNotNull<T>(this IList<T> obj, T? item)
    {
        if (item == null) return;
        obj.Add(item);
    }
}