using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TfsRetrospectiveTool
{
	public partial class Form1 : Form
	{
		private readonly Config m_config;

		private Dictionary<WorkItem, int> m_wrongAreaBugs;
		private List<WorkItem> m_leadTasks;
		private List<WorkItem> m_newFuncBugs;
		private List<WorkItem> m_regressBugs;
		private List<WorkItem> m_sdBugs;
		private List<WorkItem> m_noShipsBugs;
		private double m_ltCompletedSum;

		public Form1()
		{
			InitializeComponent();

			m_config = ConfigManager.LoadConfig();

			tfsUrlTextBox.Text = m_config.TfsUrl;
			areaPathTextBox.Text = m_config.AreaPath;
			iterationTextBox.Text = m_config.Iteration;
		}

		private void Form1FormClosing(object sender, FormClosingEventArgs e)
		{
			ConfigManager.SaveConfig(m_config);
		}

		private void HandleException(
			Exception exc,
			string caption)
		{
			var strBuilder = new StringBuilder();
			AppendExceptionString(exc, strBuilder);
			string text = strBuilder.ToString();
			using (var fileWriter = new StreamWriter(DateTime.Now.ToString("yyyy-mm-dd HH-mm-ss") + ".txt", false))
			{
				fileWriter.WriteLine(text);
			}
			Invoke(new Action(() => MessageBox.Show(text, caption)));
		}

		private void AppendExceptionString(Exception exc, StringBuilder stringBuilder)
		{
			if (exc.InnerException != null)
				AppendExceptionString(exc.InnerException, stringBuilder);
			stringBuilder.AppendLine(exc.Message);
			stringBuilder.AppendLine(exc.StackTrace);
		}

		private void ProgressReport(int percent, Label label)
		{
			//if (percent > 100)
			//	return;
			string percentStr = percent + "%";
			if (label.Text == percentStr)
				return;
			if (InvokeRequired)
				Invoke(new Action(() => label.Text = percentStr));
			else
				label.Text = percentStr;
		}

		private void ToggleMainControls(bool isEnabled)
		{
			tfsUrlTextBox.Enabled = isEnabled;
			areaPathTextBox.Enabled = isEnabled;
			iterationTextBox.Enabled = isEnabled;
			ltSearchButton.Enabled = isEnabled;
			wrongAreaBugsButton.Enabled = isEnabled;
			newFuncBugsButton.Enabled = isEnabled;
			regressBugsButton.Enabled = isEnabled;
			sdBugsButton.Enabled = isEnabled;
			noShipBugsButton.Enabled = isEnabled;
		}

		private void SaveSettingsToConfig()
		{
			m_config.TfsUrl = tfsUrlTextBox.Text;
			m_config.AreaPath = areaPathTextBox.Text;
			m_config.Iteration = iterationTextBox.Text;
		}

		private void LtSearchButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			ltPercentLabel.Text = "0%";
			ltPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => GetLeadTasks());
		}

		private void GetLeadTasks()
		{
			try
			{
				m_leadTasks = DataLoader.GetLeadTasks(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, ltPercentLabel));
				SaveSettingsToConfig();
				var ltStats = StatisticsCalculator.GetLtStats(m_leadTasks);
				m_ltCompletedSum = ltStats.Item2;
				Invoke(new Action(() =>
					{
						ltLabel.Text = m_leadTasks.Count.ToString(CultureInfo.InvariantCulture);
						ltEstimateLabel.Text = ltStats.Item1.ToString(CultureInfo.InvariantCulture);
						ltCompletedLabel.Text = ltStats.Item2.ToString(CultureInfo.InvariantCulture);
						ltPlanErrorLabel.Text = ltStats.Item3.ToString("P", CultureInfo.InvariantCulture);
						groupBox2.Enabled = true;
						ltExportButton.Enabled = true;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					ltPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void WrongAreaBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			wrongAreaBugsPercentLabel.Text = "0%";
			wrongAreaBugsPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => SearchWrongAreaPathBugs());
		}

		private void SearchWrongAreaPathBugs()
		{
			try
			{
				m_wrongAreaBugs = DataLoader.GetWrongAreaBugs(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, wrongAreaBugsPercentLabel));
				SaveSettingsToConfig();
				Invoke(new Action(() =>
				{
					wrongAreaBugsLabel.Text = m_wrongAreaBugs.Count.ToString(CultureInfo.InvariantCulture);
					fixBugsButton.Visible = m_wrongAreaBugs.Count > 0;
					if (m_wrongAreaBugs.Count == 0)
						groupBox3.Enabled = true;
					else
						wrongAreaBugsExportButton.Enabled = true;
				}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					wrongAreaBugsPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void FixBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			fixBugsButton.Enabled = false;
			fixPercentLabel.Text = "0%";
			fixPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => FixBugs());
		}

		private void FixBugs()
		{
			try
			{
				DataUploader.FixBugsAreaPaths(
					m_config.TfsUrl,
					m_wrongAreaBugs,
					x => ProgressReport(x, fixPercentLabel));
				Invoke(new Action(() =>
					{
						fixBugsButton.Visible = false;
						wrongAreaBugsLabel.Text = "???";
						groupBox3.Enabled = true;
						wrongAreaBugsExportButton.Enabled = false;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					fixBugsButton.Enabled = true;
					fixPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void NewFuncBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			newFuncBugsPercentLabel.Text = "0%";
			newFuncBugsPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => GetNewFuncBugs());
		}

		private void GetNewFuncBugs()
		{
			try
			{
				m_newFuncBugs = DataLoader.GetNewFuncBugs(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, newFuncBugsPercentLabel));
				SaveSettingsToConfig();
				var bugStats = StatisticsCalculator.GetBugStats(m_newFuncBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						newFuncBugsLabel.Text = m_newFuncBugs.Count.ToString(CultureInfo.InvariantCulture);
						newFuncBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						newFuncBugsRatioLabel.Text = bugStats.Item2.ToString("P", CultureInfo.InvariantCulture);
						groupBox4.Enabled = true;
						newFuncBugsExportButton.Enabled = true;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					newFuncBugsPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void RegressBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			regressBugsPercentLabel.Text = "0%";
			regressBugsPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => GetRegressBugs());
		}

		private void GetRegressBugs()
		{
			try
			{
				m_regressBugs = DataLoader.GetRegressBugs(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, regressBugsPercentLabel));
				SaveSettingsToConfig();
				var bugStats = StatisticsCalculator.GetBugStats(m_regressBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						regressBugsLabel.Text = m_regressBugs.Count.ToString(CultureInfo.InvariantCulture);
						regressBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						regressBugsRatioLabel.Text = bugStats.Item2.ToString("P", CultureInfo.InvariantCulture);
						groupBox5.Enabled = true;
						regressBugsExportButton.Enabled = true;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					regressBugsPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void SdBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			sdBugsPercentLabel.Text = "0%";
			sdBugsPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => GetSdBugs());
		}

		private void GetSdBugs()
		{
			try
			{
				m_sdBugs = DataLoader.GetSdBugs(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, sdBugsPercentLabel));
				SaveSettingsToConfig();
				var bugStats = StatisticsCalculator.GetBugStats(m_sdBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						sdBugsLabel.Text = m_sdBugs.Count.ToString(CultureInfo.InvariantCulture);
						sdBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						groupBox6.Enabled = true;
						sdBugsExportButton.Enabled = true;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					sdBugsPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void NoShipBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			noShipBugsPercentLabel.Text = "0%";
			noShipBugsPercentLabel.Visible = true;

			ThreadPool.QueueUserWorkItem(x => GetNoShipBugs());
		}

		private void GetNoShipBugs()
		{
			try
			{
				m_noShipsBugs = DataLoader.GetNoShipBugs(
					tfsUrlTextBox.Text,
					areaPathTextBox.Text,
					iterationTextBox.Text,
					x => ProgressReport(x, noShipBugsPercentLabel));
				SaveSettingsToConfig();
				var bugStats = StatisticsCalculator.GetBugStats(m_noShipsBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						noShipBugsLabel.Text = m_noShipsBugs.Count.ToString(CultureInfo.InvariantCulture);
						noShipBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						noShipBugsExportButton.Enabled = true;
					}));
			}
			catch (Exception e)
			{
				HandleException(e, "Error");
			}
			Invoke(new Action(() =>
				{
					noShipBugsPercentLabel.Visible = false;
					ToggleMainControls(true);
				}));
		}

		private void LtExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_leadTasks);
		}

		private void WrongAreaBugsExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_wrongAreaBugs.Keys.ToList());
		}

		private void NewFuncBugsExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_newFuncBugs);
		}

		private void RegressBugsExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_regressBugs);
		}

		private void SdBugsExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_sdBugs);
		}

		private void NoShipBugsExportButtonClick(object sender, EventArgs e)
		{
			WorkItemsToExcelExporter.Export(m_noShipsBugs);
		}
	}
}
