using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace JsonExcel
{
	/// <summary>
	/// Checks Branch-Object
	/// </summary>
	[TestClass]
	public class BranchTest
	{
		public BranchTest()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		private TestContext testContextInstance;

		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		public TestContext TestContext
		{
			get
			{
				return this.testContextInstance;
			}
			set
			{
				this.testContextInstance = value;
			}
		}

		#region Additional test attributes

		//
		// You can use the following additional attributes as you write your tests:
		//
		// Use ClassInitialize to run code before running the first test in the class
		// [ClassInitialize()]
		// public static void MyClassInitialize(TestContext testContext) { }
		//
		// Use ClassCleanup to run code after all tests in a class have run
		// [ClassCleanup()]
		// public static void MyClassCleanup() { }
		//
		// Use TestInitialize to run code before running each test
		// [TestInitialize()]
		// public void MyTestInitialize() { }
		//
		// Use TestCleanup to run code after each test has run
		// [TestCleanup()]
		// public void MyTestCleanup() { }
		//

		#endregion Additional test attributes

		/// <summary>
		/// Checks Equal-Implementation.
		/// </summary>
		[TestMethod]
		public void BranchSimpleEqualTest()
		{
			Branch originalBranch = new Branch("One");
			Assert.AreNotEqual(originalBranch, null, "Branch compared with null must be false");
			Branch anotherBranch = new Branch("Two");
			Assert.AreNotEqual(originalBranch, anotherBranch, "Two Branches with different names should be not equal");

			Branch sameAsOriginal = new Branch("One");
			Assert.AreEqual(originalBranch, sameAsOriginal, "Two Branches with equal names should be equal");
			Assert.AreNotSame(originalBranch, sameAsOriginal, "Just to be sure that the equal branches are not the same");
		}

		/// <summary>
		/// Checks Equal-Implementation with Parent-Branches.
		/// </summary>
		[TestMethod]
		public void BranchEqualWithParentTest()
		{
			Branch root = new Branch("Root");
			Branch firstBranch = new Branch("Branch_1")
			{
				Parent = root
			};
			Branch firstBranchWithoutParent = new Branch(firstBranch.Name);
			int hashWithoutParent = firstBranchWithoutParent.GetHashCode();
			Assert.AreNotEqual(firstBranch, firstBranchWithoutParent, "Branches are not equal because of Paret-Branch");
			firstBranchWithoutParent.Parent = root;
			Assert.AreEqual(firstBranch, firstBranchWithoutParent, "Branches are equal because the Paret-Branch is equal as well");
			Assert.AreNotEqual(hashWithoutParent, firstBranchWithoutParent.GetHashCode(), "HashCode must be different with Parent");
		}
	}
}