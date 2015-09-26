using System;
using System.Collections.Generic;
using System.Linq;
using TfsUtils.Accessors;

namespace TfsRetrospectiveTool
{
	internal class DataUploader
	{
		private const int LoadingPart = 50;

		internal static void FixBugsAreaPaths(
			string tfsUrl,
			Dictionary<int, int> bugsWithLinks,
			Action<int> progressReportHandler)
		{
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				var ids = bugsWithLinks.Values.Concat(bugsWithLinks.Keys).ToList();
				var items = wiqlAccessor.QueryWorkItemsByIds(
					ids,
					null,
					progressReportHandler == null
						? null
						: new Action<int>(x => progressReportHandler(x * LoadingPart / 100)));
				var dict = items.ToDictionary(i => i.Id);
				int ind = 0;
				foreach (var pair in bugsWithLinks)
				{
					string areaPath = dict[pair.Value].AreaPath;

					var bug = dict[pair.Key];
					bug.AreaPath = areaPath;

					bug.Save();

					if (progressReportHandler != null)
						progressReportHandler(LoadingPart + ind * (100 - LoadingPart) / bugsWithLinks.Count);
					++ind;
				}
			}
		}
	}
}
