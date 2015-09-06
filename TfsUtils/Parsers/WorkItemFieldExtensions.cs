using System;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsUtils.Parsers
{
	public static class WorkItemFieldExtensions
	{
		public static string Discipline(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Discipline");
		}

		public static string AssignedTo(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Assigned To");
		}

		public static int? Priority(this WorkItem workItem)
		{
			object priority = workItem.Fields.Contains("Priority")
				? workItem["Priority"]
				: workItem["Backlog Priority"];
			return priority == null ? (int?)null : (int)double.Parse(priority.ToString());
		}

		public static DateTime? StartDate(this WorkItem workItem)
		{
			return GetNullableValue<DateTime>(workItem, "Start Date");
		}

		public static DateTime? FinishDate(this WorkItem workItem)
		{
			return GetNullableValue<DateTime>(workItem, "Finish Date");
		}

		public static double? Estimate(this WorkItem workItem)
		{
			return GetNullableValue<double>(workItem, "Estimate");
		}

		public static double? Remaining(this WorkItem workItem)
		{
			return GetNullableValue<double>(workItem, "Remaining Work");
		}

		public static double? Completed(this WorkItem workItem)
		{
			return GetNullableValue<double>(workItem, "Completed Work");
		}

		public static double? ChildrenCompleted(this WorkItem workItem)
		{
			return GetNullableValue<double>(workItem, "Children Completed Work");
		}

		public static bool IsDevCompleted(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Dev Completed Agreed") == "Yes";
		}

		public static string HlaAgreementState(this WorkItem workItem)
		{
			return GetStringValue(workItem, "HLA Agreed");
		}

		public static string VisionAgreementState(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Vision Agreed");
		}

		public static string BlockingReason(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Blocking Reason");
		}

		public static string Sprint(this WorkItem workItem)
		{
			return GetStringValue(workItem, "Sprint");
		}

		private static string GetStringValue(WorkItem workItem, string fieldName)
		{
			if (!workItem.Fields.Contains(fieldName))
				return string.Empty;

			object fieldObj = workItem[fieldName];
			if (fieldObj == null)
				return string.Empty;

			return fieldObj.ToString();
		}

		private static T? GetNullableValue<T>(WorkItem workItem, string fieldName)
			where T : struct 
		{
			if (!workItem.Fields.Contains(fieldName))
				return null;
			return workItem[fieldName] as T?;
		}
	}
}
