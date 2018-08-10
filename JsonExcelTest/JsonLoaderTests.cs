using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;

namespace JsonExcel
{
	[TestClass]
	public class JsonLoaderTests
	{
		/// <summary>
		/// Check JsonLoader.FindMaxDepth to find the deepest branch in Json.
		/// </summary>
		[TestMethod]
		public void FindDepthInJsonTest()
		{
			string jsonTest = @"
{
	'l1': {
		'l2' : {
			'l3': {
				'l4' : 'Hallo'
			}
		},
		'l11' : 'Nix',
		'l12' : 'Nix2'
	},
	'l21' : {
		'l211' : 'huhu'
	}
}
";

			JObject o3 = JObject.Parse(jsonTest);
			int depth = JsonExcelLib.FindMaxDepth(o3);
			Assert.AreEqual(depth, 5, "Biggest depth of this Json should be 5");
		}

		/// <summary>
		/// Test AddNewBranch-Method to Branch.
		/// </summary>
		[TestMethod]
		public void ParseStringArrayToBranchTest()
		{
			string[][] xxx = new string[][]
				{
					new string[] {"[passenger]","[message]","[id_document]","","","","[error]",":","Cannot load document types!"},
					new string[] {"[passenger]","[message]","[save]","","","","[success]",":","Passenger data was saved!"},
					new string[] {"[passenger]","[message]","[save]","","","","[error]",":","Cannot save passenger!"},
					new string[] {"[passenger]","[message]","[save]","","","","[outward_first]",":","Outward trip's data need to be filled in!"},
					new string[] {"[passenger]","[message]","[mismatch]","","","","[any]",":","The data you scanned is different from information entered previously."},
					new string[] {"[passenger]","[message]","[mismatch]","","","","[check]",":"," Please check the details highlighted below."},
					new string[] {"[passenger]","[message]","","","","","[error]",":","Cannot load passengers!"},
					new string[] {"[passenger]","[label]","","","","","[add_owner]",":","Add owner"},
					new string[] {"[passenger]","[label]","","","","","[edit_owner]",":","Edit owner"},
					new string[] {"[passenger]","[label]","","","","","[passengers]",":","Passengers"},
					new string[] {"[passenger]","[label]","","","","","[passenger]",":","Passenger"},
					new string[] {"[passenger]","[label]","","","","","[verifyAll]",":","Verify all residents"},
					new string[] {"[passenger]","[label]","[gender]","","","","[title]",":","Gender"},
					new string[] {"[passenger]","[label]","[gender]","","","","[M]",":","M"},
					new string[] {"[passenger]","[label]","[gender]","","","","[F]",":","F"},
					new string[] {"[passenger]","[tooltip]","","","","","[form_invalid]",":","Some necessary fields in form are empty"},
					new string[] {"[passenger]","[tooltip]","","","","","[not_paid]",":","The ticket is not settled"},
					new string[] {"[passenger]","[tooltip]","","","","","[ticket_invalid]",":","Ticket expired"},
					new string[] {"[passenger]","[tooltip]","","","","","[default]",":","Reason is unknown"},
					new string[] {"[passenger]","[tooltip]","","","","","[certificateDataInvalid]",":","Some necessary fields in subsidy data are empty."},
					new string[] {"[passenger]","[subsidy]","[label]","","","","[subsidyData]",":","Subsidy data"},
					new string[] {"[passenger]","[subsidy]","[label]","","","","[certificateNumber]",":","Certificate number"},
					new string[] {"[passenger]","[subsidy]","[label]","","","","[validTo]",":","Valid to"},
					new string[] {"[passenger]","[subsidy]","[label]","","","","[searchForCertificates]",":","Search for certificates"},
					new string[] {"[passenger]","[subsidy]","[tooltip]","","","","[verifyResidentButton]",":","First name', 'First last name' and 'DNI/NIE/NP' can not be empty."},
					new string[] {"[passenger]","[subsidy]","[tooltip]","","","","[findCertificatesButton]",":","DNI/NIE/NP' can not be empty"}
				};

			Branch root = new Branch("");

			for (int i = 0; i < xxx.Length; i++)
			{
				string[] tmpLine = xxx[i].Where(x => !String.IsNullOrWhiteSpace(x)).ToArray();
				try
				{
					Branch br = root.AddNewBranch(tmpLine);
				}
				catch (Exception ex)
				{
					Assert.Fail("An Exception occoured while add new Branch to Branch: " + ex.Message);
				}
			}
		}
	}
}