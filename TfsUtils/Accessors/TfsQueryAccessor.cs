using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsUtils.Accessors
{
	public class TfsQueryAccessor : IDisposable
	{
		private readonly WorkItemStore m_workItemStore;
		private readonly TfsAccessor m_tfsAccessor;

		public TfsQueryAccessor(string tfsUrl)
		{
			m_tfsAccessor = new TfsAccessor(tfsUrl);
			m_workItemStore = m_tfsAccessor.GetWorkItemStore();
		}

		public void Dispose()
		{
			m_tfsAccessor.Dispose();
		}

		public List<WorkItem> QueryWorkItems(string queryPath)
		{
			List<string> paths = queryPath.Split(Path.DirectorySeparatorChar).ToList();
			if (paths.Count == 1)
				paths = queryPath.Split(Path.AltDirectorySeparatorChar).ToList();

			Project project = m_workItemStore.Projects[paths[0]];
			paths.RemoveAt(0);

			QueryItem item = project.QueryHierarchy;
			foreach (string path in paths)
			{
				var queryFolder = item as QueryFolder;
				item = queryFolder[path];
			}
			if (item == null)
				throw new InvalidOperationException("Invalid Query Path.");

			string queryText = m_workItemStore.GetQueryDefinition(item.Id).QueryText.Replace("@project", "'" + project.Name + "'");
			var query = new Query(m_workItemStore, queryText);
			if (query.IsLinkQuery)
			{
				var queryResults = query.RunLinkQuery();
				var workItemIds = new List<int>();
				foreach (var queryResult in queryResults)
				{
					workItemIds.Add(queryResult.SourceId);
					workItemIds.Add(queryResult.TargetId);
				}
				using (var wiqlAccessor = new TfsWiqlAccessor(m_tfsAccessor))
				{
					return wiqlAccessor.QueryWorkItemsByIds(
						workItemIds,
						null,
						null);
				}
			}
			var qResult = query.RunQuery();
			var result = new List<WorkItem>(qResult.Count);
			for (int i = 0; i < qResult.Count; i++)
			{
				result.Add(qResult[i]);
			}
			return result;
		}
	}
}
