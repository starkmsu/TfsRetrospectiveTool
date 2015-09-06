using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsRetrospectiveTool
{
	internal class WorkItemsToExcelExporter
	{
		private static readonly List<string> s_fields = new List<string>
		{
			"Id",
			"Title",
			"Estimate",
			"Completed Work",
			"Children Completed Work",
		};

		internal static void Export(IList<WorkItem> workItems)
		{
			var app = new Application();
			var workBook = app.Workbooks.Add();
			var sheet = (Worksheet)workBook.Worksheets.Item[1];

			var visibleFields = new List<string>();
			for (int i = 0; i < s_fields.Count; i++)
			{
				var field = s_fields[i];
				bool isVisible = workItems.Any(w =>
					w.Fields[field].Value != null);
				if (!isVisible)
					continue;
				sheet.Cells[1, 1 + i] = field;
				visibleFields.Add(field);
			}
			for (int ind = 0; ind < workItems.Count; ind++)
			{
				var workItem = workItems[ind];
				for (int i = 0; i < visibleFields.Count; i++)
				{
					var field = visibleFields[i];
					sheet.Cells[2+ind, 1 + i] = workItem.Fields[field].Value;
				}
			}

			app.Visible = true;
		}
	}
}
