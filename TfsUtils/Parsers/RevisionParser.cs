using System;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsUtils.Parsers
{
	public class RevisionParser
	{
		public static string GetChangeUser(Revision revision)
		{
			return revision == null ? null : revision["Changed By"].ToString();
		}

		public static DateTime GetChangedDate(Revision revision)
		{
			return DateTime.Parse(revision["Changed Date"].ToString()).Date;
		}

		public static double GetCompletedWorkChange(Revision revision)
		{
			double curVal = GetCompletedWork(revision);
			if (curVal <= 0)
				return 0;

			Revision prevRevision = revision.Index == 0
				? null
				: revision.WorkItem.Revisions[revision.Index - 1];
			double prevVal = GetCompletedWork(prevRevision);
			return curVal - prevVal;
		}

		public static double GetCompletedWork(Revision revision)
		{
			if (revision == null)
				return 0;

			object obj = revision["Completed Work"];
			return obj == null ? 0 : (double)obj;
		}

		public static int GetChahngesetId(Revision revision)
		{
			int prevLinksCount = revision.Index == 0
				? 0
				: revision.WorkItem.Revisions[revision.Index - 1].Links.Count;
			if (prevLinksCount == revision.Links.Count)
				return 0;

			ExternalLink lastChangeset = null;
			for (int i = 0; i < revision.Links.Count; i++)
			{
				var link = revision.Links[i];
				if (link.BaseType == BaseLinkType.ExternalLink)
				{
					var extLink = link as ExternalLink;
					if (extLink.ArtifactLinkType.Name != "Fixed in Changeset")
						continue;
					lastChangeset = extLink;
				}
				if (lastChangeset != null && link.BaseType != BaseLinkType.ExternalLink)
					break;
			}

			if (lastChangeset == null)
				return 0;

			string url = lastChangeset.LinkedArtifactUri;
			int lastInd = url.LastIndexOf('/');
			return int.Parse(url.Substring(lastInd + 1, url.Length - 1 - lastInd));
		}
	}
}
