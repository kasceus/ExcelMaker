using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMaker.Model_Classes
{
	/// <summary>
	/// The row of data held for render into the excel document
	/// </summary>
	internal class Row
	{
		/// <summary>
		/// Id for the record
		/// </summary>
		internal string recordId { get; set; }
		/// <summary>
		/// List of column data stored for the row
		/// </summary>
		internal List<RowData> data { get; set; }
	}
}
