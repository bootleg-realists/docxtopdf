using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BootlegRealists.Reporting.Enumeration;

// TODO: handle sectPr w:docGrid w:linePitch (how many lines per page)
// TODO: handle tblLook $17.7.6 (conditional formatting), it will define firstrow/firstcolumn/...etc styles in styles.xml

namespace BootlegRealists.Reporting.Extension;

/// <summary>
/// This class contains open xml element extension methods.
/// </summary>
public static class OpenXmlElementExtension
{
	/// <summary>
	/// Gets the first element from the element list of the given object.
	/// </summary>
	/// <typeparam name="T">Target element type</typeparam>
	/// <param name="obj">The object to act on</param>
	/// <returns>The element or default otherwise</returns>
	public static T? GetFirstElement<T>(this OpenXmlElement obj) where T : OpenXmlElement
	{
		return obj.Elements().FirstOrDefault(x => x.GetType() == typeof(T)) as T;
	}

	/// <summary>
	/// Gets the first descendant from the children of the given object.
	/// </summary>
	/// <typeparam name="T">Target element type</typeparam>
	/// <param name="obj">The object to act on</param>
	/// <returns>The descendant or default otherwise</returns>
	public static T? GetFirstDescendant<T>(this OpenXmlElement obj) where T : OpenXmlElement
	{
		return obj.Descendants<T>().FirstOrDefault();
	}

	/// <summary>
	/// Copy the attributes from the given object to the given destination object.
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <param name="dest">The destination of attributes copy process.</param>
	public static void CopyAttributesTo(this OpenXmlElement obj, OpenXmlElement? dest)
	{
		if (dest == null)
			return;

		dest.ClearAllAttributes();
		dest.SetAttributes(obj.GetAttributes());
	}

	/// <summary>
	/// Appends the attributes from the given object to the given destination object.
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <param name="dest">The destination of attributes copy process.</param>
	public static void AppendAttributesTo(this OpenXmlElement obj, OpenXmlElement dest)
	{
		dest.SetAttributes(obj.GetAttributes());
	}

	/// <summary>
	/// Gets the main document part for the given object
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <returns>The main document part</returns>
	public static OpenXmlPart? GetMainDocumentPart(this OpenXmlElement obj)
	{
		var result = GetRootPart(obj);
		if (result == null) return null;
		if (result is MainDocumentPart) return result;
		var parents = result.GetParentParts();
		return parents.FirstOrDefault(x => x.GetType() == typeof(MainDocumentPart));
	}

	/// <summary>
	/// Get the nearest applied element (e.g. RunFonts, Languages) in the style hierarchy. It searches upstream from
	/// current obj til reach the top of the hierarchy if no found.
	/// </summary>
	/// <typeparam name="T">Target element type (e.g. RunFonts.GetType()).</typeparam>
	/// <param name="obj">The OpenXmlElement to search from.</param>
	/// <returns>Return found element or default otherwise.</returns>
	public static T? GetEffectiveElement<T>(this OpenXmlElement obj) where T : OpenXmlElement
	{
		var result = default(T);

		return obj switch
		{
			null => result,
			Paragraph paragraph => GetEffectiveElementParagraph<T>(paragraph),
			Table table => GetEffectiveElementTable<T>(table),
			TableRow row => GetEffectiveElementTableRow<T>(row),
			TableCell cell => GetEffectiveElementTableCell<T>(cell),
			Style style => GetEffectiveElementStyle<T>(style),
			OpenXmlCompositeElement run => GetEffectiveElementRun<T>(run),
			_ => obj.GetFirstDescendant<T>()
		};
	}

	/// <summary>
	/// Gets the root part for the given object
	/// </summary>
	/// <param name="obj">The object to act on</param>
	/// <returns>The root part</returns>
	static OpenXmlPart? GetRootPart(this OpenXmlElement obj)
	{
		OpenXmlPart? result = obj.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart;
		if (result != null) return result;
		result = obj.Ancestors<Footer>().FirstOrDefault()?.FooterPart;
		return result ?? obj.Ancestors<Header>().FirstOrDefault()?.HeaderPart;
	}

	static T? GetEffectiveElementRun<T>(OpenXmlElement obj) where T : OpenXmlElement
	{
		var list = new List<T>();
		// Run.RunProperties > Run.RunProperties.rStyle > Paragraph.ParagraphProperties.pStyle (> default style > docDefaults)
		// ( ): done in paragraph level
		var runProperties = obj.GetFirstElement<RunProperties>();
		T? result;
		if (runProperties != null)
		{
			result = runProperties.GetFirstDescendant<T>();
			list.AddNotNull(result);

			// If has rStyle, go through rStyle before go on.
			// Use getAppliedStyleElement() is because it will go over all the basedOn styles.
			if (runProperties.RunStyle != null)
			{
				result = GetEffectiveElementStyle<T>(runProperties.RunStyle.GetStyleById());
				list.AddNotNull(result);
			}
		}

		if (obj.Parent is Paragraph { ParagraphProperties.ParagraphStyleId: not null } paragraph) // parent paragraph's pStyle
		{
			result = GetEffectiveElementStyle<T>(paragraph.ParagraphProperties.ParagraphStyleId.GetStyleById());
			list.AddNotNull(result);
		}

		// default run style
		if (obj.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart)
			return null;
		result = GetEffectiveElementStyle<T>(mainDocumentPart.GetDefaultStyle(DefaultStyleType.Character));
		list.AddNotNull(result);
		result = GetEffectiveElementStyle<T>(mainDocumentPart.GetDefaultStyle(DefaultStyleType.Paragraph));
		list.AddNotNull(result);

		var docDefaults = mainDocumentPart.StyleDefinitionsPart?.Styles?.Descendants<DocDefaults>().FirstOrDefault();
		result = docDefaults?.RunPropertiesDefault?.GetFirstDescendant<T>();
		list.AddNotNull(result);

		result = MergeProps(list);
		return result;
	}

	static T? GetEffectiveElementParagraph<T>(OpenXmlElement obj) where T : OpenXmlElement
	{
		var list = new List<T>();
		// Paragraph.ParagraphProperties > Paragraph.ParagraphProperties.pStyle > default style > docDefaults
		var paragraphProperties = obj.GetFirstElement<ParagraphProperties>();
		T? result;
		if (paragraphProperties != null)
		{
			result = paragraphProperties.GetFirstDescendant<T>();
			list.AddNotNull(result);

			// If has pStyle, go through pStyle before go on.
			// Use getAppliedStyleElement() is because it will go over the whole Style hierarchy.
			if (paragraphProperties.ParagraphStyleId != null)
			{
				result = GetEffectiveElementStyle<T>(paragraphProperties.ParagraphStyleId.GetStyleById());
				list.AddNotNull(result);
			}
		}

		if (obj.Parent is TableCell)
		{
			var table = obj.Ancestors<Table>().FirstOrDefault();
			if (table != null)
			{
				result = GetEffectiveElementTable<T>(table);
				list.AddNotNull(result);
			}
		}

		if (obj.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart)
			return null;
		result = GetEffectiveElementStyle<T>(mainDocumentPart.GetDefaultStyle(DefaultStyleType.Paragraph));
		list.AddNotNull(result);

		var docDefaults = mainDocumentPart.StyleDefinitionsPart?.Styles?.Descendants<DocDefaults>().FirstOrDefault();
		result = docDefaults?.GetFirstDescendant<T>();
		list.AddNotNull(result);

		result = MergeProps(list);
		return result;
	}

	static T? GetEffectiveElementTable<T>(OpenXmlElement obj) where T : OpenXmlElement
	{
		var list = new List<T>();
		// Table.TableProperties > Table.TableProperties.tblStyle > default style
		var tableProperties = obj.GetFirstElement<TableProperties>();
		T? result;
		if (tableProperties != null)
		{
			result = tableProperties.GetFirstDescendant<T>();
			if (result != null)
			{
				list.Add(result);
			}
			// If has tblStyle, go through tblStyle before go on.
			// Use getAppliedStyleElement() is because it will go over the whole Style hierarchy.
			else if (tableProperties.TableStyle != null)
			{
				result = GetEffectiveElementStyle<T>(tableProperties.TableStyle.GetStyleById());
				list.AddNotNull(result);
			}
		}
		if (obj.GetMainDocumentPart() is not MainDocumentPart mainDocumentPart)
			return null;
		result = GetEffectiveElementStyle<T>(mainDocumentPart.GetDefaultStyle(DefaultStyleType.Table));
		list.AddNotNull(result);

		result = MergeProps(list);
		return result;
	}

	static T? GetEffectiveElementTableRow<T>(OpenXmlElement obj) where T : OpenXmlElement
	{
		var list = new List<T>();
		T? result;
		// TableRow.TableRowProperties > Table.TableProperties.tblStyle (> default style)
		// ( ): done in Table level
		var rowProperties = obj.GetFirstElement<TableRowProperties>();
		if (rowProperties != null)
		{
			result = rowProperties.GetFirstDescendant<T>();
			list.AddNotNull(result);
		}

		var root = obj;
		for (var i = 0; i < 1 && root != null; i++) root = root.Parent;

		if (root is Table table)
		{
			result = GetEffectiveElementTable<T>(table);
			list.AddNotNull(result);
		}
		result = MergeProps(list);
		return result;
	}

	static T? GetEffectiveElementTableCell<T>(OpenXmlElement obj) where T : OpenXmlElement
	{
		var list = new List<T>();
		T? result;
		// TableCell.TableCellProperties > Table.TableProperties.tblStyle (> default style)
		// ( ): done in Table level
		var cellProperties = obj.GetFirstElement<TableCellProperties>();
		if (cellProperties != null)
		{
			result = cellProperties.GetFirstDescendant<T>();
			list.AddNotNull(result);
		}

		var root = obj;
		for (var i = 0; i < 2 && root != null; i++) root = root.Parent;

		if (root is Table table)
		{
			result = GetEffectiveElementTable<T>(table);
			list.AddNotNull(result);
		}
		result = MergeProps(list);
		return result;
	}

	static T? GetEffectiveElementStyle<T>(Style? obj) where T : OpenXmlElement
	{
		if (obj == null) return default;
		var result = obj.GetFirstDescendant<T>();
		if (result != null) return result;
		if (obj.BasedOn != null)
			result = GetEffectiveElementStyle<T>(obj.BasedOn.GetStyleById());
		return result;
	}

	/// <summary>
	/// Merges a list of items into one item by filling the empty properties from the current with the next one from the list.
	/// </summary>
	/// <typeparam name="T">Type of the list</typeparam>
	/// <returns>The merged item</returns>
	static T? MergeProps<T>(IList<T> list) where T : class
	{
		if (list.Count == 0) return null;
		var result = list[0];
		var props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
		for (var i = 1; i < list.Count; i++)
		{
			var emptyProps = GetEmptyProps(props, result);
			foreach (var emptyProp in emptyProps)
			{
				var val = emptyProp.GetValue(list[i]);
				if (val != null)
					emptyProp.SetValue(result, val);
			}
		}
		return result;
	}

	/// <summary>
	/// Gets a list of properties which are empty (null)
	/// </summary>
	/// <param name="props">List of properties</param>
	/// <param name="obj">Object to check for empty properties</param>
	/// <returns>List of empty properties</returns>
	static IEnumerable<PropertyInfo> GetEmptyProps(IEnumerable<PropertyInfo> props, object obj)
	{
		var result = new List<PropertyInfo>();
		foreach (var prop in props)
		{
			if (prop.GetValue(obj) == null)
				result.Add(prop);
		}
		return result;
	}
}
