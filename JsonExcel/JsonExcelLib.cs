using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace JsonExcel
{
	public static class JsonExcelLib
	{
		public static JObject LoadJsonToJObject(string filename)
		{
			JObject jsonObj = null;
			if (!String.IsNullOrWhiteSpace(filename) && File.Exists(filename))
			{
				try
				{
					using (StreamReader fileStream = File.OpenText(filename))
					using (JsonTextReader reader = new JsonTextReader(fileStream))
					{
						jsonObj = (JObject)JToken.ReadFrom(reader);
					}
				}
				catch (Exception ex)
				{
					throw JsonExcelException.LoadFileException(ex);
				}
			}
			else
			{
				throw JsonExcelException.LoadFileNoFileFoundException();
			}

			if (jsonObj == null)
			{
				throw JsonExcelException.ParseFileException();
			}
			return jsonObj;
		}

		/// <summary>
		/// Find the maximum depth of branches in Json.
		/// </summary>
		/// <param name="jsonObject"></param>
		/// <returns></returns>
		public static int FindMaxDepth(this JObject jsonObject)
		{
			int bigTmp = 0;
			if (jsonObject == null)
				return bigTmp;

			foreach (KeyValuePair<string, JToken> tokensFromJOject in jsonObject)
			{
				int tmpDepth = FindMaxDepth(tokensFromJOject.Value, 0);
				bigTmp = Math.Max(bigTmp, tmpDepth);
			}
			return bigTmp;
		}

		private static int FindMaxDepth(JToken token, int currentDepth)
		{
			if (token == null)
				return currentDepth;

			int bigTmp = currentDepth;
			if (token.Children().Any())
			{
				foreach (var property in token.Children())
				{
					int tmpDepth = FindMaxDepth(property, currentDepth + 1);
					bigTmp = Math.Max(bigTmp, tmpDepth);
				}
			}
			else
			{
				bigTmp = 0;
			}
			return bigTmp;
		}

		/// <summary>
		/// Json-Representation of Branch.
		/// </summary>
		/// <param name="br"></param>
		/// <returns></returns>
		public static string ToJsonString(this Branch br)
		{
			return GetJsonStringFromBranch(br, 0);
		}

		private static string GetJsonStringFromBranch(Branch br, int depth)
		{
			if (br == null || String.IsNullOrWhiteSpace(br.Value))
				return "";

			StringBuilder sb = new StringBuilder(new String(' ', depth * 2));
			if (!String.IsNullOrWhiteSpace(br.Name))
			{
				sb.Append($"\"{br.Name}\"");
				sb.Append(" : ");
			}

			int max = br.Children.Count();
			if (max > 0)
			{
				sb.AppendLine("{");
				int count = 0;
				foreach (Branch nb in br.Children)
				{
					sb.Append(GetJsonStringFromBranch(nb, depth + 1));
					sb.AppendLine((++count < max) ? "," : "");
				}
				sb.Append(' ', depth * 2);
				sb.Append("}");
			}
			else
			{
				string v = br.Value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\'", "\\\'").Replace("\'", "\\\'");
				sb.Append($"\"{v}\"");
			}
			return sb.ToString();
		}

		/// <summary>
		/// Add a new Branch that was created from StringArray.
		/// </summary>
		/// <param name="rootBranch"></param>
		/// <param name="stringArrThatConvertsToNewBranch"></param>
		/// <returns></returns>
		public static Branch AddNewBranch(this Branch rootBranch, string[] stringArrThatConvertsToNewBranch)
		{
			return AddNewBranchFromLineToRoot(rootBranch, stringArrThatConvertsToNewBranch, 0);
		}

		private static Branch AddNewBranchFromLineToRoot(Branch rootBranch, string[] stringArrThatConvertsToNewBranch, int idx)
		{
			if (rootBranch == null || stringArrThatConvertsToNewBranch == null || stringArrThatConvertsToNewBranch.Count() < 3 || !stringArrThatConvertsToNewBranch.Where(x => x.Trim().Equals(":")).Any())
			{
				throw new Exception("Data could not parse to new Branch");
			}

			Branch retVal = rootBranch;
			if (idx >= stringArrThatConvertsToNewBranch.Length - 1)
				return retVal;

			string linePart = stringArrThatConvertsToNewBranch[idx];
			if (linePart[0].Equals('['))
				linePart = linePart.Substring(1, linePart.Length - 2);

			if (linePart.Trim().Equals(":"))
			{
				retVal.Value = stringArrThatConvertsToNewBranch[idx + 1];
				return retVal;
			}

			IEnumerable<Branch> hasChild = retVal.Children.Where(x => x.Name.Equals(linePart));
			if (hasChild.Any())
			{
				AddNewBranchFromLineToRoot(hasChild.First(), stringArrThatConvertsToNewBranch, idx + 1);
			}
			else
			{
				Branch newBranch = new Branch(linePart)
				{
					Parent = retVal
				};
				AddNewBranchFromLineToRoot(newBranch, stringArrThatConvertsToNewBranch, idx + 1);
			}
			return retVal;
		}
	}
}