using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonExcel
{
	[ComVisible(true)]
	public class JsonRibbon : ExcelRibbon
	{
		#region CONST-Values

		private static readonly Color STRUCTUREPART_FONTCOLOR = Color.Black;
		private static readonly Color STRUCTUREPART_BACKCOLOR = Color.LightGray;
		private static readonly Color KEYPART_FONTCOLOR = Color.Black;
		private static readonly Color KEYPART_BACKCOLOR = Color.LightGreen;
		private static readonly Color VALUEPART_FONTCOLOR = Color.Red;
		private static readonly Color VALUEPART_BACKCOLOR = Color.Yellow;

		#endregion CONST-Values

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
              <button id='LoadJsonFile' label='Load JSON-File' onAction='OnButtonPressed_LoadJsonFile'/>
              <button id='SaveAsJson' label='Save as JSON-File' onAction='OnButtonPressed_SaveAsJson'/>
              <button id='Parse' label='Parse to JSON' onAction='OnButtonPressed_ParseToJson'/>
              <button id='Info' label='Info' onAction='OnButtonPressed_ShowInfo'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
		}

		public void OnButtonPressed_ShowInfo(IRibbonControl control)
		{
			MessageBox.Show("JsonExcel-AddIn (c) Tjark Onnen", "JsonExcel-AddIn-Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
				root = ParseSheetToBranch();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: Could not parse Data. Original error: " + ex.Message);
				return;
			}
			string jsonText = root.ToJsonString();
			if (String.IsNullOrWhiteSpace(jsonText))
			{
				MessageBox.Show("Error: No Data to parse.");
				return;
			}
			try
			{
				JObject jsonObj = JObject.Parse(jsonText);
				this._gapToShowValuesInSameColumn = jsonObj.FindMaxDepth();
				int rowNumber = 1;
				foreach (KeyValuePair<string, JToken> jsonToken in jsonObj)
				{
					rowNumber = ShowTokenInSheet(jsonToken.Value, rowNumber, 2, new List<string> { jsonToken.Key });
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: Could not Show Json in Excel: " + ex.Message);
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
					root = ParseSheetToBranch();
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error: Could not parse Data. File not saved. Original error: " + ex.Message);
					return;
				}
				string jsonText = root.ToJsonString();
				if (String.IsNullOrWhiteSpace(jsonText))
				{
					MessageBox.Show("Error: No Data to parse. No File was saved.");
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
					MessageBox.Show("Error: Could not save file to disk. Original error: " + ex.Message);
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
				RestoreDirectory = true
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
					MessageBox.Show(ex.Message);
					return;
				}

				this._gapToShowValuesInSameColumn = jsonObj.FindMaxDepth();

				int rowNumber = 1;
				foreach (KeyValuePair<string, JToken> jsonToken in jsonObj)
				{
					rowNumber = ShowTokenInSheet(jsonToken.Value, rowNumber, 2, new List<string> { jsonToken.Key });
				}
			}
		}

		private Branch ParseSheetToBranch()
		{
			Branch root = new Branch("");

			for (int i = 1; i <= 5000; i++)
			{
				Excel.Range line = this._activeWorksheet.Cells.Range[this._activeWorksheet.Cells[i, 1], this._activeWorksheet.Cells[i, 50]];
				string[] strArrayFromExcelRow = line.Cells.Cast<Excel.Range>().Select(x => (string)x.Text).Where(x => !(String.IsNullOrWhiteSpace(x))).ToArray<string>();
				if (!strArrayFromExcelRow.Any())
					break;
				try
				{
					Branch br = root.AddNewBranch(strArrayFromExcelRow);
				}
				catch (Exception ex)
				{
					throw JsonExcelException.ParseBranchException($"Error: Line {i} could not parsed", ex);
				}
			}
			return root;
		}

		private int ShowTokenInSheet(JToken token, int rowNumber, int depth, List<string> jsonStructure)
		{
			if (token == null)
				return rowNumber;
			string childNode = "";
			if (token is JValue)
			{
				childNode = ShowRowInSheet(token, rowNumber, depth, jsonStructure);
				rowNumber++;
			}
			else if (token is JObject obj)
			{
				foreach (var property in obj.Properties())
				{
					childNode = property.Name;
					if (jsonStructure.Count() < depth)
					{
						jsonStructure.Add(childNode);
					}
					else
					{
						jsonStructure[depth - 1] = childNode;
					}
					rowNumber = ShowTokenInSheet(property.Value, rowNumber, depth + 1, jsonStructure);
				}
			}

			return rowNumber;
		}

		private string ShowRowInSheet(JToken token, int lineCount, int depth, List<string> keys)
		{
			Show_1_StructurePart_1(lineCount, depth, keys);
			Show_2_KeyPart(lineCount, depth, keys);
			return Show_3_ValuePart(token, lineCount);
		}

		private string Show_3_ValuePart(JToken token, int lineCount)
		{
			string childNode = token.ToString();
			Excel.Range ran = this._activeWorksheet.Cells[lineCount, this._gapToShowValuesInSameColumn + 2];
			ran.Value2 = childNode;
			ran.Interior.Color = VALUEPART_BACKCOLOR;
			ran.Font.Color = VALUEPART_FONTCOLOR;
			return childNode;
		}

		private void Show_2_KeyPart(int lineCount, int depth, List<string> keys)
		{
			Excel.Range ran = this._activeWorksheet.Cells[lineCount, this._gapToShowValuesInSameColumn];
			ran.Value2 = $"[{keys[depth - 2]}]";
			ran.Interior.Color = Color.LightGreen;
			ran.Font.Color = Color.Black;

			ran = this._activeWorksheet.Cells[lineCount, this._gapToShowValuesInSameColumn + 1];
			ran.Value2 = ":";
		}

		private void Show_1_StructurePart_1(int lineCount, int depth, List<string> keys)
		{
			for (int i = 0; i < depth - 2; i++)
			{
				Excel.Range ran = this._activeWorksheet.Cells[lineCount, i + 1];
				ran.Value2 = $"[{keys[i]}]";
				ran.Interior.Color = STRUCTUREPART_BACKCOLOR;
				ran.Font.Color = STRUCTUREPART_FONTCOLOR;
			}
			Excel.Range clearRange = this._activeWorksheet.Range[this._activeWorksheet.Cells[lineCount, depth - 1], this._activeWorksheet.Cells[lineCount, this._gapToShowValuesInSameColumn - 1]];
			clearRange.Clear();
		}
	}
}