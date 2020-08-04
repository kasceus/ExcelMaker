using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelMaker.Model_Classes
{
	/// <summary>
	/// The worksheet is the basic building block of an excel document
	/// </summary>
	public partial class WorkSheet
	{
		private XLWorkbook wbd = new XLWorkbook();
		private const string invalidCharsRegex = @"[/\\*'?[\]:]+";
		private string sheetName { get; set; }
		internal IXLWorksheet worksheet { get; set; }
		/// <summary>
		/// Used to create a worksheet that has multiple data tables on a single sheet
		/// <para>Pass single DataTable for single table per sheet</para>
		/// </summary>
		/// <param name="tables">List of Data Tables to add to the sheet</param>
		/// <param name="sheetName">Name to display for the sheet</param>
		public WorkSheet(List<DataTable> tables, string sheetName = "Sheet")
		{
			string safeName = Regex.Replace(sheetName, invalidCharsRegex, " ")
									.Replace("  ", " ")
									.Trim();
			this.sheetName = safeName;
			worksheet = convertToWorksheet(tables);
		}
		/// <summary>
		/// Used to create a worksheet 
		/// <para>Pass a list of DataTables to create a sheet with multiple tables per sheet</para>
		/// </summary>
		/// <param name="table">DataTable to add to the sheet.</param>
		/// <param name="sheetName">Name for the sheet.  If unspecified, then the dataTable's name will be used (if any)</param>
		public WorkSheet(DataTable table, string sheetName = "Sheet")
		{
			if (sheetName == "Sheet" && string.IsNullOrEmpty(table.TableName) == false)
			{
				sheetName = table.TableName;
			}
			string safeName = Regex.Replace(sheetName, invalidCharsRegex, " ")
									.Replace("  ", " ")
									.Trim();

			this.sheetName = safeName;
			worksheet = convertToWorksheet(table);
		}
		/// <summary>
		/// Use to create a worksheet from generic type data.
		/// </summary>
		/// <param name="data">Generic IEnumerable data to convert into an excel sheet</param>
		/// <param name="sheetName">Name for the sheet. If unspecified, sheet name will default to Type name of the generic object passed in</param>
		public WorkSheet(IEnumerable<object> data, string sheetName = "Sheet")
		{
			string safeName = Regex.Replace(sheetName, invalidCharsRegex, " ")
									.Replace("  ", " ")
									.Trim();
			this.sheetName = safeName;
			worksheet = convertToWorksheet(data);
		}
		private IXLWorksheet convertToWorksheet(IEnumerable<object> data)
		{
			DataTable dt = new DataTable();
			Type type = data.GetType().GetTypeInfo().GenericTypeArguments[0];
			if (sheetName == "Sheet")
			{
				sheetName = type.Name ?? "Sheet";
			}
			string fQN = type.AssemblyQualifiedName;//full assembly name required here to get the proper type castings
			List<Heads> heads = dt.MakeHeaders(fQN);//columns already set for the data table.  pass the ordered heads to the body maker method
			dt.MakeRows(data, heads);
			dt.TableName = sheetName;
			IXLWorksheet sheet = Methods.MakeWorksheet(dt);
			sheet.Name = sheetName;

			return sheet;
		}
		private IXLWorksheet convertToWorksheet(List<DataTable> tables)
		{
			IXLWorksheet worksheet = Methods.MakeWorksheet(tables);
			worksheet.Name = sheetName;
			return worksheet;
		}
		private IXLWorksheet convertToWorksheet(DataTable table)
		{
			table.TableName = table.TableName ?? sheetName;
			IXLWorksheet worksheet = Methods.MakeWorksheet(table);

			worksheet.Name = sheetName;

			return worksheet;
		}
	}
}
