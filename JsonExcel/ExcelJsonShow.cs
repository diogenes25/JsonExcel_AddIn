using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonExcel
{
	public static class ExcelJsonShow
	{
		#region CONST-Values

		private static readonly Color STRUCTUREPART_FONTCOLOR = Color.Black;
		private static readonly Color STRUCTUREPART_BACKCOLOR = Color.LightGray;
		private static readonly Color KEYPART_FONTCOLOR = Color.Black;
		private static readonly Color KEYPART_BACKCOLOR = Color.LightGreen;
		private static readonly Color VALUEPART_FONTCOLOR = Color.Red;
		private static readonly Color VALUEPART_BACKCOLOR = Color.Yellow;

		#endregion CONST-Values

		public static Branch ParseSheetToBranch(Excel.Worksheet ws)
		{
			Branch root = new Branch("");

			for (int i = 1; i <= 5000; i++)
			{
				Microsoft.Office.Interop.Excel.Range line = ws.Cells.Range[ws.Cells[i, 1], ws.Cells[i, 50]];
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

		public static int ShowTokenInSheet(JToken token, int rowNumber, int depth, List<string> jsonStructure, Excel.Worksheet ws, int gap)
		{
			if (token == null)
				return rowNumber;
			string childNode = "";
			if (token is JValue)
			{
				childNode = ShowRowInSheet(token, rowNumber, depth, jsonStructure, ws, gap);
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
					rowNumber = ShowTokenInSheet(property.Value, rowNumber, depth + 1, jsonStructure, ws, gap);
				}
			}

			return rowNumber;
		}

		private static string ShowRowInSheet(JToken token, int lineCount, int depth, List<string> keys, Excel.Worksheet ws, int gap)
		{
			Show_1_StructurePart_1(lineCount, depth, keys, ws, gap);
			Show_2_KeyPart(lineCount, depth, keys, ws, gap);
			return Show_3_ValuePart(token, lineCount, ws, gap);
		}

		private static void Show_1_StructurePart_1(int lineCount, int depth, List<string> keys, Excel.Worksheet ws, int gap)
		{
			for (int i = 0; i < depth - 2; i++)
			{
				Excel.Range ran = (Excel.Range)ws.Cells[lineCount, i + 1];
				ran.Value2 = $"[{keys[i]}]";
				ran.Interior.Color = STRUCTUREPART_BACKCOLOR;
				ran.Font.Color = STRUCTUREPART_FONTCOLOR;
			}
			if (depth < gap)
			{
				Excel.Range clearRange = ws.Range[ws.Cells[lineCount, depth - 1], ws.Cells[lineCount, gap - 1]];
				clearRange.Clear();
			}
		}

		private static void Show_2_KeyPart(int lineCount, int depth, List<string> keys, Excel.Worksheet ws, int gap)
		{
			Excel.Range ran = (Excel.Range)ws.Cells[lineCount, gap];
			ran.Value2 = $"[{keys[depth - 2]}]";
			ran.Interior.Color = KEYPART_BACKCOLOR;
			ran.Font.Color = KEYPART_FONTCOLOR;

			ran = (Excel.Range)ws.Cells[lineCount, gap + 1];
			ran.Value2 = ":";
		}

		private static string Show_3_ValuePart(JToken token, int lineCount, Excel.Worksheet ws, int gap)
		{
			Excel.Range ran = (Excel.Range)ws.Cells[lineCount, gap + 2];
			//string childNode = token.ToString();
			switch (token.Type)
			{
				case JTokenType.Integer:
				case JTokenType.Float:
				case JTokenType.String:
					ran.Value2 = token.ToString();
					break;

				case JTokenType.Boolean:
					ran.Value2 = (bool)token ? "true" : "false";
					break;

				case JTokenType.Array:
					ran.Value2 = GetStringFromArray((JArray)token);
					break;

				default:
					ran.Value2 = token.ToString();
					break;
			}
			//ran.Value2 = childNode;
			ran.Interior.Color = VALUEPART_BACKCOLOR;
			ran.Font.Color = VALUEPART_FONTCOLOR;
			return String.Empty;
		}

		private static string GetStringFromArray(JArray array)
		{
			if (array.Count == 0)
				return "";

			bool isOneLine =
				array.Children()
					 .All(t => t.Type == JTokenType.Integer ||
							   t.Type == JTokenType.Float ||
							   t.Type == JTokenType.Boolean);
			if (isOneLine)
			{
				return string.Join(", ",
								   array.Children().Select(c => c.ToString(Formatting.None)).ToArray());
			}
			else
			{
				return string.Join("\n",
								   array.Children().Select(c => c.ToString(Formatting.None)).ToArray());
			}
		}

	}
}
