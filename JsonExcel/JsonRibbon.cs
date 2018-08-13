using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonExcel
{
	[ComVisible(true)]
	public class JsonRibbon : ExcelRibbon
	{
	
		private Excel.Application _app;
		private Excel.Worksheet _activeWorksheet;
		private string _lastPath = "c:\\";
		private int _gapToShowValuesInSameColumn = 1;

		public override string GetCustomUI(string RibbonID)
		{
			return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
	<ribbon>
		<tabs>
			<tab id='JsonTab' label='JSON'>
				<group id='jsonGrp' label='JSON File'>
					<button id='LoadJsonFile' label='Import JSON' onAction='OnButtonPressed_LoadJsonFile' size='large' visible='true' imageMso='ImportXmlFile' />
					<button id='SaveAsJson' label='Export as JSON' onAction='OnButtonPressed_SaveAsJson' size='large' visible='true' imageMso='ExportTextFile' />
					<button id='Parse' label='Parse' onAction='OnButtonPressed_ParseToJson' size='large' visible='true' imageMso='FileCompatibilityChecker' />
					<button id='Info' label='Info' onAction='OnButtonPressed_ShowInfo' />
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>";
		}

		private static T GetAssemblyInfo<T>() where T : class
		{
			return (T)Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(T), false).FirstOrDefault();
		}

		public void OnButtonPressed_ShowInfo(IRibbonControl control)
		{
			MessageBox.Show(
				   string.Format("{0}\n{1}\n{2}",
					   GetAssemblyInfo<AssemblyTitleAttribute>().Title,
					   GetAssemblyInfo<AssemblyFileVersionAttribute>().Version,
					   GetAssemblyInfo<AssemblyCopyrightAttribute>().Copyright),
				   "ExcelJsonTable",
				   MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		/// <summary>
		/// Parse current Excel-Sheet and reprint this on current Excel-Sheet.
		/// </summary>
		/// <param name="control"></param>
		public void OnButtonPressed_ParseToJson(IRibbonControl control)
		{
			this._app = (Excel.Application)ExcelDnaUtil.Application;
			this._activeWorksheet = (Excel.Worksheet)this._app.ActiveWorkbook.ActiveSheet;

			Branch root = null;
			try
			{
				root = ExcelJsonShow.ParseSheetToBranch(this._activeWorksheet);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: Could not parse Data. Original error: " + ex.Message, "Parse", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			string jsonText = root.ToJsonString();
			if (String.IsNullOrWhiteSpace(jsonText))
			{
				MessageBox.Show("Error: No Data to parse.", "Parse", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			try
			{
				JObject jsonObj = JObject.Parse(jsonText);
				this._gapToShowValuesInSameColumn = jsonObj.FindMaxDepth();
				int rowNumber = 1;
				foreach (KeyValuePair<string, JToken> jsonToken in jsonObj)
				{
					rowNumber = ExcelJsonShow.ShowTokenInSheet(jsonToken.Value, rowNumber, 2, new List<string> { jsonToken.Key }, this._activeWorksheet, this._gapToShowValuesInSameColumn);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: Could not Show Json in Excel: " + ex.Message, "Parse", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
		}

		/// <summary>
		/// Save current excel-Sheet as Json-File.
		/// </summary>
		/// <param name="control"></param>
		public void OnButtonPressed_SaveAsJson(IRibbonControl control)
		{
			this._app = (Excel.Application)ExcelDnaUtil.Application;
			this._activeWorksheet = (Excel.Worksheet)this._app.ActiveWorkbook.ActiveSheet;

			SaveFileDialog saveFileDialog = new SaveFileDialog
			{
				InitialDirectory = this._lastPath,
				Filter = "Json files (*.json)|*.json|txt files (*.txt)|*.txt|All files (*.*)|*.*",
				FilterIndex = 1,
			};

			if (saveFileDialog.ShowDialog() == DialogResult.OK)
			{
				Branch root = null;
				try
				{
					root = ExcelJsonShow.ParseSheetToBranch(this._activeWorksheet);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error: Could not parse Data. File not saved. Original error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (root == null || !root.Children.Any())
				{
					MessageBox.Show("No Data", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				string jsonText = root.ToJsonString();
				if (String.IsNullOrWhiteSpace(jsonText))
				{
					MessageBox.Show("Error: No Data to parse. No File was saved.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				try
				{
					JObject o = JObject.Parse(jsonText);
					File.WriteAllText(saveFileDialog.FileName, o.ToString());
					this._lastPath = Path.GetDirectoryName(saveFileDialog.FileName);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error: Could not save file to disk. Original error: " + ex.Message, "Export", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		/// <summary>
		/// Load Json-File and show there Data in current Excel-Sheet.
		/// </summary>
		/// <param name="control"></param>
		public void OnButtonPressed_LoadJsonFile(IRibbonControl control)
		{
			OpenFileDialog openFileDialog1 = new OpenFileDialog
			{
				InitialDirectory = this._lastPath,
				Filter = "Json files (*.json)|*.json|txt files (*.txt)|*.txt|All files (*.*)|*.*",
				FilterIndex = 1,
				RestoreDirectory = true,
				Multiselect = false,
				Title = "Select a file to import"
			};

			this._app = (Excel.Application)ExcelDnaUtil.Application;
			this._activeWorksheet = (Excel.Worksheet)this._app.ActiveWorkbook.ActiveSheet;

			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				JObject jsonObj = null;
				try
				{
					jsonObj = JsonExcelLib.LoadJsonToJObject(openFileDialog1.FileName);
					this._lastPath = Path.GetDirectoryName(openFileDialog1.FileName);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				this._gapToShowValuesInSameColumn = jsonObj.FindMaxDepth();

				int rowNumber = 1;
				foreach (KeyValuePair<string, JToken> jsonToken in jsonObj)
				{
					rowNumber = ExcelJsonShow.ShowTokenInSheet(jsonToken.Value, rowNumber, 2, new List<string> { jsonToken.Key }, this._activeWorksheet, this._gapToShowValuesInSameColumn);
				}
			}
		}

	
	}
}