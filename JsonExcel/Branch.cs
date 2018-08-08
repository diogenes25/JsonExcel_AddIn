using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JsonExcel
{
	/// <summary>
	/// Branch (a part of the Json-Tree-Structure)
	/// </summary>
	public class Branch
	{
		private string _name = "";
		private Branch _parent = null;

		/// <summary>
		/// Name of Json-Property
		/// </summary>
		public string Name
		{
			get { return this._name; }
			private set
			{
				if (!String.IsNullOrWhiteSpace(value) && value.Count() > 1 && value[0].Equals('['))
					this._name = value.Substring(1, value.Length - 2);
				else
					this._name = value;
			}
		}

		/// <summary>
		/// Value of Property.
		/// </summary>
		/// <remarks>
		/// If this value is set no children should be set.
		/// </remarks>
		public string Value { get; set; }

		/// <summary>
		/// Parent Branch
		/// </summary>
		/// <remarks>
		/// If a Parent is set, this Branch will be set as a child to the Parent-Branch.
		/// </remarks>
		public Branch Parent
		{
			get { return this._parent; }
			set
			{
				if (this._parent != null)
				{
					this._parent.Children.Remove(this);
				}

				this._parent = value;
				if (this._parent != null)
					this._parent.Children.Add(this);
			}
		}

		/// <summary>
		/// Children-Branches of this branch.
		/// </summary>
		/// <remarks>
		/// this.Value should not be set.
		/// </remarks>
		public List<Branch> Children { get; } = new List<Branch>();

		/// <summary>
		/// Branch (Property of a Json-File)
		/// </summary>
		/// <param name="name"></param>
		public Branch(string name)
		{
			this.Name = name;
		}

		public override string ToString()
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(this.Name);
			if (String.IsNullOrWhiteSpace(this.Value))
			{
				foreach (Branch c in this.Children)
				{
					sb.Append(c.ToString());
				}
			}
			else
			{
				sb.AppendLine(this.Value);
			}
			return sb.ToString();
		}

		/// <summary>
		/// Check equal.
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			if (obj is Branch)
			{
				bool parentEqual = this._parent == null ? (((Branch)obj).Parent == null) : this._parent.Equals(((Branch)obj).Parent);
				return parentEqual && this.Name.Equals(((Branch)obj).Name);
			}
			else
				return false;
		}

		/// <summary>
		/// Get HashCode.
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			int hash = this.Parent == null ? 0 : this._parent.Name.GetHashCode();
			return hash ^ this._name.GetHashCode();
		}
	}
}