using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsUtils.Parsers
{
	public class WorkItemParser
	{
		public static Dictionary<string, List<int>> ParseLinks(WorkItem workItem)
		{
			var result = new Dictionary<string, List<int>>();

			for (int j = 0; j < workItem.Links.Count; j++)
			{
				var link = workItem.Links[j];
				if (link.BaseType != BaseLinkType.RelatedLink)
					continue;

				var relLink = (RelatedLink)link;
				string linkName = relLink.LinkTypeEnd.Name;
				if (!result.ContainsKey(linkName))
					result.Add(linkName, new List<int>());
				result[linkName].Add(relLink.RelatedWorkItemId);
			}

			return result;
		}

		public static List<int> GetRelationsByType(WorkItem workItem, string relationType)
		{
			var result = new List<int>();

			for (int j = 0; j < workItem.Links.Count; j++)
			{
				var link = workItem.Links[j];
				if (link.BaseType != BaseLinkType.RelatedLink)
					continue;

				var relLink = (RelatedLink)link;
				string linkName = relLink.LinkTypeEnd.Name;
				if (linkName != relationType)
					continue;

				result.Add(relLink.RelatedWorkItemId);
			}

			return result;
		}
	}
}
