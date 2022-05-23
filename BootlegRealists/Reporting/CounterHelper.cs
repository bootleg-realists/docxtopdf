using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BootlegRealists.Reporting;

internal class CounterHelper
{
	readonly Numbering numbering;
	readonly NumberingCounter numberingCounter = new();

	public CounterHelper()
	{
		numbering = new Numbering();
	}

	public CounterHelper(WordprocessingDocument doc)
	{
		var numb = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
		numbering = numb ?? new Numbering();

		// set all ilevel's start value for all abstractNums
		foreach (var abstractNum in numbering.Descendants<AbstractNum>())
		{
			var abstractNumId = abstractNum.AbstractNumberId?.Value ?? 0;
			foreach (var level in abstractNum.Descendants<Level>())
			{
				if (level.LevelIndex != (object?)null && level.StartNumberingValue != null)
					numberingCounter.SetStart(abstractNumId, level.LevelIndex.Value, level.StartNumberingValue?.Val ?? 0);
			}
		}
	}

	/// <summary>
	/// Get paragraph's numbering level object by searching abstractNums with the numbering level ID.
	/// </summary>
	/// <param name="paragraph"></param>
	/// <returns></returns>
	public Level? GetLevel(Paragraph paragraph)
	{
		Level? ret = null;

		var pgpr = paragraph.Elements<ParagraphProperties>().FirstOrDefault();
		if (pgpr == null)
			return ret;

		// direct numbering
		var numPr = pgpr.Elements<NumberingProperties>().FirstOrDefault();
		if (numPr?.NumberingId == null)
			return ret;

		int numId = numPr.NumberingId.Val!;
		if (numId <= 0)
			return ret;

		int? ilvl = null;
		string? refStyleName = null;

		if (numPr.NumberingLevelReference != null)
		{
			// ilvl included in NumberingProperties
			ilvl = numPr.NumberingLevelReference.Val!;
		}
		else
		{
			// doesn't have ilvl in NumberingProperties, search by referenced style name
			var st = pgpr.Elements<Style>().FirstOrDefault();
			if (st?.StyleName != null) refStyleName = st.StyleName.Val!;
		}

		// find abstractNumId by numId
		var numInstance = numbering.Elements<NumberingInstance>().FirstOrDefault(c => c.NumberID?.Value == numId);
		if (numInstance == null)
			return ret;

		{
			// find abstractNum by abstractNumId
			var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(c =>
				c.AbstractNumberId?.Value == numInstance.AbstractNumId?.Val?.Value);
			if (abstractNum == null)
				return ret;

			{
				if (ilvl != null) // search by ilvl
					ret = abstractNum.Elements<Level>().FirstOrDefault(c => c.LevelIndex?.Value == ilvl);
				else if (refStyleName != null) // search by matching referenced style name
					ret = abstractNum.Elements<Level>().FirstOrDefault(c => c.ParagraphStyleIdInLevel != null && c.ParagraphStyleIdInLevel.Val == refStyleName);
			}
		}

		return ret;
	}

	/// <summary>
	/// Get a list of numbering value from level-0 to level-ilvl. Call this method will
	/// 1. increase the numbering value of level-ilvl by one automatically
	/// 2. restart all the levels larger than level-ilvl
	/// </summary>
	/// <param name="abstractNumId"></param>
	/// <param name="ilvl"></param>
	/// <returns></returns>
	public List<int> GetCurrent(int abstractNumId, int ilvl)
	{
		return numberingCounter.GetCurrent(abstractNumId, ilvl);
	}

	#region NumberingCounter

	class NumberingCounter
	{
		/// <summary>
		/// Store abstractNums, the key is abstractNumId.
		/// </summary>
		readonly Dictionary<int, List<LevelCounter>> abstractNums = new();

		public void SetStart(int abstractNumId, int ilvl, int start)
		{
			if (start < 0) return;

			if (abstractNums.ContainsKey(abstractNumId))
			{
				var lc = abstractNums[abstractNumId].Find(c => c.Level == ilvl);
				if (lc != null)
					lc.Start = start;
				else
					abstractNums[abstractNumId].Add(new LevelCounter(ilvl, start));
			}
			else
			{
				abstractNums[abstractNumId] = new List<LevelCounter>
				{
					new(ilvl, start)
				};
			}
		}

		/// <summary>
		/// Restart the numbering value of ilvl, and all the levels larger than ilvl will be restarted as well.
		/// </summary>
		/// <param name="abstractNumId"></param>
		/// <param name="ilvl"></param>
		void Restart(int abstractNumId, int ilvl)
		{
			if (!abstractNums.ContainsKey(abstractNumId))
				return;

			while (true)
			{
				var lc = abstractNums[abstractNumId].Find(c => c.Level == ilvl);
				if (lc != null)
				{
					lc.Restart();
					ilvl++;
				}
				else
				{
					break;
				}
			}
		}

		/// <summary>
		/// Get a list of numbering value from level-0 to level-ilvl. Call this method will
		/// 1. increase the numbering value of level-ilvl by one automatically
		/// 2. all the levels larger than level-ilvl will be restarted
		/// </summary>
		/// <param name="abstractNumId"></param>
		/// <param name="ilvl"></param>
		/// <returns></returns>
		public List<int> GetCurrent(int abstractNumId, int ilvl)
		{
			var ret = new List<int>();
			if (!abstractNums.ContainsKey(abstractNumId))
				return ret;

			LevelCounter? lc;
			for (var i = 0; i < ilvl; i++)
			{
				// get the numbering value from the levels smaller than ilvl
				lc = abstractNums[abstractNumId].Find(c => c.Level == i);
				if (lc != null) ret.Add(lc.CurrentStatic);
			}

			lc = abstractNums[abstractNumId].Find(c => c.Level == ilvl);
			if (lc == null)
				return ret;

			// get the numbering value from ilvl
			ret.Add(lc.Current);
			Restart(abstractNumId, ilvl + 1); // all the levels larger than ilvl should restart

			return ret;
		}

		class LevelCounter
		{
			int start;

			public LevelCounter(int ilvl, int start)
			{
				Level = ilvl;
				Start = start;
			}

			/// <summary>
			/// Get iLevel.
			/// </summary>
			public int Level { get; }

			/// <summary>
			/// Get current numbering and increase by one.
			/// </summary>
			public int Current => ++CurrentStatic;

			/// <summary>
			/// Get current numbering but without changing it.
			/// </summary>
			public int CurrentStatic { get; private set; }

			/// <summary>
			/// Set start numbering (must be not a negative value).
			/// </summary>
			public int Start
			{
				set
				{
					if (value > 0) start = CurrentStatic = value - 1;
				}
			}

			/// <summary>
			/// Reset current numbering to start value.
			/// </summary>
			public void Restart()
			{
				CurrentStatic = start;
			}
		}
	}

	#endregion
}
