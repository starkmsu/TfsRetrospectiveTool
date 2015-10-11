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
	public partial class MainForm : Form
	{
		private const string ZeroPercents = "0%";
		private const string UnknownCount = "???";

		private readonly Config m_config;

		private Dictionary<int, int> m_wrongAreaBugs;
		private List<WorkItem> m_leadTasks;
		private List<WorkItem> m_newFuncBugs;
		private List<WorkItem> m_regressBugs;
		private List<WorkItem> m_sdBugs;
		private List<WorkItem> m_noShipsBugs;
		private double m_ltCompletedSum;

		public MainForm()
		{
			InitializeComponent();

			m_config = ConfigManager.LoadConfig();

			tfsUrlTextBox.Text = m_config.TfsUrl;
			areaPathComboBox.Text = m_config.AreaPath;
			if (m_config.AllAreaPaths.Count == 0
				&& !string.IsNullOrEmpty(m_config.AreaPath))
				m_config.AllAreaPaths.Add(m_config.AreaPath);
			areaPathComboBox.DataSource = m_config.AllAreaPaths;
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
			areaPathComboBox.Enabled = isEnabled;
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
			m_config.AreaPath = areaPathComboBox.Text;
			m_config.AllAreaPaths = areaPathComboBox.Items.Cast<string>().ToList();
			m_config.Iteration = iterationTextBox.Text;
		}

		private void LtSearchButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			ltPercentLabel.Text = ZeroPercents;
			ltPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			var areaSet = new HashSet<string>();
			var allAreas = areaPathComboBox.Items
				.Cast<string>()
				.ToList();
			foreach (string path in allAreas)
			{
				areaSet.Add(path);
			}
			if (!areaSet.Contains(areaPath))
			{
				allAreas.Add(areaPath);
				areaPathComboBox.DataSource = allAreas;
				areaPathComboBox.Text = areaPath;
			}

			ThreadPool.QueueUserWorkItem(x => GetLeadTasks(areaPath));
		}

		private void GetLeadTasks(string areaPath)
		{
			try
			{
				m_leadTasks = DataLoader.GetLeadTasks(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, ltPercentLabel));
				var ltStats = StatisticsCalculator.GetLtStats(m_leadTasks);
				m_ltCompletedSum = ltStats.Item2;
				Invoke(new Action(() =>
					{
						SaveSettingsToConfig();
						ltLabel.Text = m_leadTasks.Count.ToString(CultureInfo.InvariantCulture);
						ltEstimateLabel.Text = ltStats.Item1.ToString(CultureInfo.InvariantCulture);
						ltCompletedLabel.Text = ltStats.Item2.ToString(CultureInfo.InvariantCulture);
						ltPlanErrorLabel.Text = ltStats.Item3.ToString("P", CultureInfo.InvariantCulture);
						ltExportButton.Enabled = true;
						groupBox2.Enabled = true;
						wrongAreaBugsLabel.Text = UnknownCount;
						wrongAreaBugsExportButton.Enabled = false;
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

			wrongAreaBugsPercentLabel.Text = ZeroPercents;
			wrongAreaBugsPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			ThreadPool.QueueUserWorkItem(x => SearchWrongAreaPathBugs(areaPath));
		}

		private void SearchWrongAreaPathBugs(string areaPath)
		{
			try
			{
				m_wrongAreaBugs = DataLoader.GetWrongAreaBugs(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, wrongAreaBugsPercentLabel));
				Invoke(new Action(() =>
				{
					wrongAreaBugsLabel.Text = m_wrongAreaBugs.Count.ToString(CultureInfo.InvariantCulture);
					fixBugsButton.Visible = m_wrongAreaBugs.Count > 0;
					if (m_wrongAreaBugs.Count != 0)
						wrongAreaBugsExportButton.Enabled = true;
					else
					{
						groupBox3.Enabled = true;
						newFuncBugsLabel.Text = UnknownCount;
						newFuncBugsCompletedLabel.Text = UnknownCount;
						newFuncBugsRatioLabel.Text = UnknownCount;
						newFuncBugsExportButton.Enabled = false;
					}
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
			fixPercentLabel.Text = ZeroPercents;
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
						wrongAreaBugsLabel.Text = @"0";
						wrongAreaBugsExportButton.Enabled = false;
						groupBox3.Enabled = true;
						newFuncBugsLabel.Text = UnknownCount;
						newFuncBugsCompletedLabel.Text = UnknownCount;
						newFuncBugsRatioLabel.Text = UnknownCount;
						newFuncBugsExportButton.Enabled = false;
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
					wrongAreaBugsLabel.Text = UnknownCount;
					ToggleMainControls(true);
				}));
		}

		private void NewFuncBugsButtonClick(object sender, EventArgs e)
		{
			ToggleMainControls(false);

			newFuncBugsPercentLabel.Text = ZeroPercents;
			newFuncBugsPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			ThreadPool.QueueUserWorkItem(x => GetNewFuncBugs(areaPath));
		}

		private void GetNewFuncBugs(string areaPath)
		{
			try
			{
				m_newFuncBugs = DataLoader.GetNewFuncBugs(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, newFuncBugsPercentLabel));
				var bugStats = StatisticsCalculator.GetBugStats(m_newFuncBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						newFuncBugsLabel.Text = m_newFuncBugs.Count.ToString(CultureInfo.InvariantCulture);
						newFuncBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						newFuncBugsRatioLabel.Text = bugStats.Item2;
						newFuncBugsExportButton.Enabled = true;
						groupBox4.Enabled = true;
						regressBugsLabel.Text = UnknownCount;
						regressBugsCompletedLabel.Text = UnknownCount;
						regressBugsRatioLabel.Text = UnknownCount;
						regressBugsExportButton.Enabled = false;
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

			regressBugsPercentLabel.Text = ZeroPercents;
			regressBugsPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			ThreadPool.QueueUserWorkItem(x => GetRegressBugs(areaPath));
		}

		private void GetRegressBugs(string areaPath)
		{
			try
			{
				m_regressBugs = DataLoader.GetRegressBugs(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, regressBugsPercentLabel));
				var bugStats = StatisticsCalculator.GetBugStats(m_regressBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						regressBugsLabel.Text = m_regressBugs.Count.ToString(CultureInfo.InvariantCulture);
						regressBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						regressBugsRatioLabel.Text = bugStats.Item2;
						regressBugsExportButton.Enabled = true;
						groupBox5.Enabled = true;
						sdBugsLabel.Text = UnknownCount;
						sdBugsCompletedLabel.Text = UnknownCount;
						sdBugsExportButton.Enabled = false;
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

			sdBugsPercentLabel.Text = ZeroPercents;
			sdBugsPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			ThreadPool.QueueUserWorkItem(x => GetSdBugs(areaPath));
		}

		private void GetSdBugs(string areaPath)
		{
			try
			{
				m_sdBugs = DataLoader.GetSdBugs(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, sdBugsPercentLabel));
				var bugStats = StatisticsCalculator.GetBugStats(m_sdBugs, m_ltCompletedSum);
				Invoke(new Action(() =>
					{
						sdBugsLabel.Text = m_sdBugs.Count.ToString(CultureInfo.InvariantCulture);
						sdBugsCompletedLabel.Text = bugStats.Item1.ToString(CultureInfo.InvariantCulture);
						sdBugsExportButton.Enabled = true;
						groupBox6.Enabled = true;
						noShipBugsLabel.Text = UnknownCount;
						noShipBugsCompletedLabel.Text = UnknownCount;
						noShipBugsExportButton.Enabled = false;
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

			noShipBugsPercentLabel.Text = ZeroPercents;
			noShipBugsPercentLabel.Visible = true;

			string areaPath = areaPathComboBox.Text;

			ThreadPool.QueueUserWorkItem(x => GetNoShipBugs(areaPath));
		}

		private void GetNoShipBugs(string areaPath)
		{
			try
			{
				m_noShipsBugs = DataLoader.GetNoShipBugs(
					tfsUrlTextBox.Text,
					areaPath,
					iterationTextBox.Text,
					x => ProgressReport(x, noShipBugsPercentLabel));
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
			var items = DataLoader.GetWorkItemsByIds(tfsUrlTextBox.Text, m_wrongAreaBugs.Keys.ToList());
			WorkItemsToExcelExporter.Export(items);
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

		private void AreaPathComboBoxTextUpdate(object sender, EventArgs e)
		{
			groupBox2.Enabled = false;
			groupBox3.Enabled = false;
			groupBox4.Enabled = false;
			groupBox5.Enabled = false;
			groupBox6.Enabled = false;
			ltLabel.Text = UnknownCount;
			ltEstimateLabel.Text = UnknownCount;
			ltCompletedLabel.Text = UnknownCount;
			ltPlanErrorLabel.Text = UnknownCount;
			ltExportButton.Enabled = false;
		}
	}
}
