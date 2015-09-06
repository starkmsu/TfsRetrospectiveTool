using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using TfsUtils.Parsers;

namespace TfsRetrospectiveTool
{
	internal class StatisticsCalculator
	{
		internal static Tuple<double, double, double> GetLtStats(List<WorkItem> leadtasks)
		{
			if (leadtasks.Count == 0)
				return new Tuple<double, double, double>(0, 0, 0);

			double estimateSum = 0;
			double complectedSum = 0;
			foreach (var leadtask in leadtasks)
			{
				double? est = leadtask.Estimate();
				estimateSum += est.HasValue ? est.Value : 0;

				double? completed = leadtask.Completed();
				complectedSum += completed.HasValue ? completed.Value : 0;

				double? chCompleted = leadtask.ChildrenCompleted();
				complectedSum += chCompleted.HasValue ? chCompleted.Value : 0;
			}

			return new Tuple<double, double, double>(
				estimateSum,
				complectedSum,
				(estimateSum - complectedSum) / estimateSum);
		}

		internal static Tuple<double, double> GetBugStats(List<WorkItem> bugs, double compareSum)
		{
			double complectedSum = 0;
			foreach (var bug in bugs)
			{
				double? completed = bug.Completed();
				complectedSum += completed.HasValue ? completed.Value : 0;
			}

			return new Tuple<double, double>(
				complectedSum,
				complectedSum / compareSum);
		}
	}
}
