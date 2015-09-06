using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsUtils.Accessors
{
	public class TfsWiqlAccessor : IDisposable
	{
		private readonly WorkItemStore m_workItemStore;
		private readonly bool m_shouldDisposeOfAccessor;
		private readonly TfsAccessor m_tfsAccessor;

		public TfsWiqlAccessor(string tfsUrl)
		{
			m_tfsAccessor = new TfsAccessor(tfsUrl);
			m_shouldDisposeOfAccessor = true;
			m_workItemStore = m_tfsAccessor.GetWorkItemStore();
		}

		public TfsWiqlAccessor(TfsAccessor tfsAccessor)
		{
			m_tfsAccessor = tfsAccessor;
			m_shouldDisposeOfAccessor = false;
			m_workItemStore = m_tfsAccessor.GetWorkItemStore();
		}

		public void Dispose()
		{
			if (m_shouldDisposeOfAccessor)
				m_tfsAccessor.Dispose();
		}

		public string AddUsersConditions(string wiqlString, List<string> users)
		{
			if (users == null || users.Count == 0)
				return wiqlString;

			return wiqlString + GenerateUsersConditions(users);
		}

		public List<WorkItem> QueryWorkItemsByIds(
			ICollection<int> ids,
			string orderByString,
			Action<int> progressReportHandler)
		{
			string queryStr = "SELECT * FROM WorkItems WHERE [System.Id] IN (@ids)";
			if (!string.IsNullOrEmpty(orderByString))
				queryStr = queryStr + " " + orderByString;

			var paramValues = new Dictionary<string, object>();

			var complexParamValues = new Dictionary<string, List<object>>
			{
				{"ids", ids.Select(i => (object)i).ToList()}
			};

			return QueryWorkItems(
				queryStr,
				paramValues,
				complexParamValues,
				progressReportHandler);
		}

		public List<WorkItem> QueryWorkItems(
			string wiqlString,
			Dictionary<string, object> paramValues,
			Dictionary<string, List<object>> complexParamValues,
			Action<int> progressReportHandler)
		{
			if (complexParamValues != null && complexParamValues.Count > 0)
				wiqlString = UpdateParams(wiqlString, paramValues, complexParamValues);

			var query = new Query(m_workItemStore, wiqlString, paramValues ?? new Dictionary<string, object>(0));

			var queryResult = query.RunQuery();
			var result = new List<WorkItem>(queryResult.Count);
			for (int i = 0; i < queryResult.Count; i++)
			{
				var wi = queryResult[i];
				wi.SyncToLatest();
				result.Add(wi);
				if (progressReportHandler != null)
					progressReportHandler(i * 100 / queryResult.Count);
			}
			return result;
		}

		public Dictionary<int, List<int>> QueryIdsFromLinks(
			string wiqlString,
			Dictionary<string, object> paramValues,
			Dictionary<string, List<object>> complexParamValues,
			Action<int> progressReportHandler)
		{
			if (complexParamValues != null && complexParamValues.Count > 0)
				wiqlString = UpdateParams(wiqlString, paramValues, complexParamValues);

			var query = new Query(m_workItemStore, wiqlString, paramValues);

			var queryResult = query.RunLinkQuery();
			var result = new Dictionary<int, List<int>>();
			for (int i = 0; i < queryResult.Length; i++)
			{
				var link = queryResult[i];
				if (link.SourceId == 0)
					result.Add(link.TargetId, new List<int>());
				else
					result[link.SourceId].Add(link.TargetId);
				if (progressReportHandler != null)
					progressReportHandler(i * 100 / queryResult.Length);
			}
			return result;
		}

		private string UpdateParams(
			string wiqlString,
			Dictionary<string, object> paramValues,
			Dictionary<string, List<object>> complexParamValues)
		{
			var strBuilder = new StringBuilder(wiqlString);
			foreach (var complexParamValue in complexParamValues)
			{
				if (paramValues.ContainsKey(complexParamValue.Key))
					continue;

				string complexParamKey = "@" + complexParamValue.Key;
				if (wiqlString.IndexOf(complexParamKey, StringComparison.OrdinalIgnoreCase) == -1
					|| complexParamValue.Value.Count == 0)
					continue;

				var paramStrBuilder = new StringBuilder();
				for (int i = 0; i < complexParamValue.Value.Count; i++)
				{
					if (i > 0)
						paramStrBuilder.Append(',');
					paramStrBuilder.Append('@');
					string paramKey = complexParamValue.Key + i;
					paramStrBuilder.Append(paramKey);

					paramValues.Add(paramKey, complexParamValue.Value[i]);
				}
				strBuilder.Replace(complexParamKey, paramStrBuilder.ToString());
			}
			return strBuilder.ToString();
		}

		private string GenerateUsersConditions(List<string> users)
		{
			var strBuilder = new StringBuilder();
			strBuilder.Append(" AND (");
			for (int i = 0; i < users.Count; i++)
			{
				if (i > 0)
					strBuilder.Append(" OR ");
				strBuilder.Append("[System.ChangedBy] EVER '");
				strBuilder.Append(users[i]);
				strBuilder.Append('\'');
			}
			strBuilder.Append(')');
			return strBuilder.ToString();
		}
	}
}
