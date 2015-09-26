using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using TfsUtils.Accessors;

namespace TfsRetrospectiveTool
{
	internal class DataLoader
	{
		internal static List<WorkItem> GetWorkItemsByIds(string tfsUrl, List<int> ids)
		{
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				return wiqlAccessor.QueryWorkItemsByIds(
					ids,
					null,
					null);
			}
		}

		internal static Dictionary<int, int> GetWrongAreaBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				var notOurBugs = GetWrongAreaPathBugs(
					wiqlAccessor,
					areaPath,
					iterarion,
					false,
					progressReportHandler);

				var ourBugs = GetWrongAreaPathBugs(
					wiqlAccessor,
					areaPath,
					iterarion,
					true,
					progressReportHandler);

				foreach (var pair in ourBugs)
				{
					notOurBugs.Add(pair.Key, pair.Value);
				}

				return notOurBugs;
			}
		}

		private static Dictionary<int, int> GetWrongAreaPathBugs(
			TfsWiqlAccessor wiqlAccessor,
			string areaPath,
			string iterarion,
			bool ourBugs,
			Action<int> progressReportHandler)
		{
			var strBuilder = new StringBuilder();
			strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
			strBuilder.Append(" WHERE Source.[System.WorkItemType] = 'Bug'");
			strBuilder.Append(" AND Source.[System.AreaPath] "
				+ (ourBugs ? "NOT UNDER '" : "UNDER '")
				+ areaPath
				+ "'");
			strBuilder.Append(" AND Source.[System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" AND Target.[System.WorkItemType] = 'Ship'");
			strBuilder.Append(" AND Target.[System.AreaPath] "
				+ (ourBugs ? "UNDER '" : "NOT UNDER '")
				+ areaPath
				+ "'");
			strBuilder.Append(" AND Target.[System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" MODE (MustContain)");

			var bugsWithOtherShips = wiqlAccessor.QueryIdsFromLinks(
				strBuilder.ToString(),
				null,
				null,
				progressReportHandler);

			if (bugsWithOtherShips.Count == 0)
				return new Dictionary<int, int>();

			strBuilder.Clear();
			strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
			strBuilder.Append(" WHERE Source.[System.Id] IN (" + string.Join(",", bugsWithOtherShips.Keys) + ")");
			strBuilder.Append(" AND Target.[System.WorkItemType] = 'Ship'");
			strBuilder.Append(" AND Target.[System.AreaPath] "
				+ (ourBugs ? "NOT UNDER '" : "UNDER '")
				+ areaPath
				+ "'");
			strBuilder.Append(" AND Target.[System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" MODE (DoesNotContain)");

			var result = wiqlAccessor.QueryIdsFromLinks(
				strBuilder.ToString(),
				null,
				null,
				progressReportHandler);

			return result.ToDictionary(i => i.Key, i => bugsWithOtherShips[i.Key].First());
		}

		internal static List<WorkItem> GetLeadTasks(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			var strBuilder = new StringBuilder();
			strBuilder.Append("SELECT * FROM WorkItems");
			strBuilder.Append(" WHERE [System.WorkItemType] = 'LeadTask'");
			strBuilder.Append(" AND [Microsoft.VSTS.Common.Discipline] = 'Development'");
			strBuilder.Append(" AND [System.AreaPath] UNDER '" + areaPath + "'");
			strBuilder.Append(" AND [System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" AND [System.Reason] <> 'Canceled'");
			strBuilder.Append(" AND [Children Completed Work] > 0");

			List<WorkItem> result;
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				result = wiqlAccessor.QueryWorkItems(
					strBuilder.ToString(),
					null,
					null,
					progressReportHandler);
			}
			return result;
		}

		internal static List<WorkItem> GetNewFuncBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			return GetBugs(
				tfsUrl,
				areaPath,
				iterarion,
				true,
				false,
				progressReportHandler);
		}

		internal static List<WorkItem> GetRegressBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			return GetBugs(
				tfsUrl,
				areaPath,
				iterarion,
				false,
				false,
				progressReportHandler);
		}

		internal static List<WorkItem> GetSdBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			return GetBugs(
				tfsUrl,
				areaPath,
				iterarion,
				false,
				true,
				progressReportHandler);
		}

		private static List<WorkItem> GetBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			bool newFunc,
			bool sd,
			Action<int> progressReportHandler)
		{
			var strBuilder = new StringBuilder();
			strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
			strBuilder.Append(" WHERE Source.[System.WorkItemType] = 'Bug'");
			strBuilder.Append(" AND Source.[System.AreaPath] UNDER '" + areaPath + "'");
			strBuilder.Append(" AND Source.[System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" AND Source.[System.State] <> 'Proposed'");
			strBuilder.Append(" AND Source.[System.State] <> 'Active'");
			if (!sd)
			{
				strBuilder.Append(" AND Source.[Found On Level] <> '6 Product Testing'");
				strBuilder.Append(" AND Source.[Regress] = '" + (newFunc ? "No" : "Yes") + "'");
			}
			strBuilder.Append(" AND Source.[Service Desk] "+ (sd ? "=" : "<>") + " 'Yes'");
			strBuilder.Append(" AND Source.[System.Reason] NOT IN ('Rejected', 'Deferred', 'Duplicate', 'Cannot Reproduce', 'Converted to Requirement')");
			strBuilder.Append(" AND Target.[System.WorkItemType] = 'LeadTask'");
			strBuilder.Append(" AND Target.[Microsoft.VSTS.Common.Discipline] = 'Development'");
			strBuilder.Append(" AND [System.Links.LinkType] = 'Child'");
			strBuilder.Append(" MODE (DoesNotContain)");

			List<WorkItem> result;
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				var ids = wiqlAccessor.QueryIdsFromLinks(
					strBuilder.ToString(),
					null,
					null,
					null);

				if (ids.Count == 0)
					return new List<WorkItem>(0);

				strBuilder.Clear();
				strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
				strBuilder.Append(" WHERE Source.[System.Id] IN (" + string.Join(",", ids.Keys) + ")");
				strBuilder.Append(" AND Target.[System.WorkItemType] = 'Ship'");
				strBuilder.Append(" AND Target.[System.AreaPath] UNDER '" + areaPath + "'");
				strBuilder.Append(" AND Target.[System.IterationPath] UNDER '" + iterarion + "'");
				strBuilder.Append(" MODE (MustContain)");

				ids = wiqlAccessor.QueryIdsFromLinks(
					strBuilder.ToString(),
					null,
					null,
					null);

				if (ids.Count == 0)
					return new List<WorkItem>(0);

				result = wiqlAccessor.QueryWorkItemsByIds(
					ids.Keys,
					"ORDER BY [Completed Work] DESC",
					progressReportHandler);
			}
			return result;
		}

		internal static List<WorkItem> GetNoShipBugs(
			string tfsUrl,
			string areaPath,
			string iterarion,
			Action<int> progressReportHandler)
		{
			var strBuilder = new StringBuilder();
			strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
			strBuilder.Append(" WHERE Source.[System.WorkItemType] = 'Bug'");
			strBuilder.Append(" AND Source.[System.AreaPath] UNDER '" + areaPath + "'");
			strBuilder.Append(" AND Source.[System.IterationPath] UNDER '" + iterarion + "'");
			strBuilder.Append(" AND Source.[System.State] <> 'Proposed'");
			strBuilder.Append(" AND Source.[System.State] <> 'Active'");
			strBuilder.Append(" AND Source.[System.Reason] <> 'Converted to Requirement'");
			strBuilder.Append(" AND Source.[Completed Work] > 0");
			strBuilder.Append(" AND Target.[System.WorkItemType] = 'Ship'");
			strBuilder.Append(" MODE (DoesNotContain)");

			List<WorkItem> result;
			using (var wiqlAccessor = new TfsWiqlAccessor(tfsUrl))
			{
				var ids = wiqlAccessor.QueryIdsFromLinks(
					strBuilder.ToString(),
					null,
					null,
					null);

				if (ids.Count == 0)
					return new List<WorkItem>(0);

				strBuilder.Clear();
				strBuilder.Append("SELECT [System.Id] FROM WorkItemLinks");
				strBuilder.Append(" WHERE Source.[System.Id] IN (" + string.Join(",", ids.Keys) + ")");
				strBuilder.Append(" AND [System.Links.LinkType] = 'Child'");
				strBuilder.Append(" AND Target.[System.WorkItemType] = 'LeadTask'");
				strBuilder.Append(" AND Target.[Microsoft.VSTS.Common.Discipline] = 'Development'");
				strBuilder.Append(" MODE (DoesNotContain)");

				ids = wiqlAccessor.QueryIdsFromLinks(
					strBuilder.ToString(),
					null,
					null,
					null);

				if (ids.Count == 0)
					return new List<WorkItem>(0);

				result = wiqlAccessor.QueryWorkItemsByIds(
					ids.Keys,
					"ORDER BY [Completed Work] DESC",
					progressReportHandler);
			}
			return result;
		}
	}
}
