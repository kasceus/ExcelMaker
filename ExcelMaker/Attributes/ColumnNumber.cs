using System;
using System.Web.Mvc;

namespace ExcelMaker.Attributes
{
	/// <summary>
	/// Set what column number this will be in the excel document
	/// </summary>
	[AttributeUsage(AttributeTargets.Property)]
	public class ColumnNumberAttribute : FilterAttribute
	{
		/// <summary>
		/// Column Number for the property
		/// </summary>
		public int columnNumber => col;

		private int col { get; set; }
		/// <summary>
		///  Set what column number this will be in the excel document
		/// </summary>
		/// <param name="number"></param>
		public ColumnNumberAttribute(int number)
		{
			if (number < 0)
			{
				throw new ArgumentException(message: "The number supplied must be 0 or more.", paramName: nameof(number));
			}
			col = number;
		}
	}
}
