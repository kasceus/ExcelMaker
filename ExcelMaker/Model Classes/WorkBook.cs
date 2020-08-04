using ClosedXML.Excel;
using System.Collections.Generic;

namespace ExcelMaker.Model_Classes
{
	/// <summary>
	/// The workbook contains all settings and sheets required for exporting an excel document
	/// <para>The easiest way to export excel documents would be to work directly with WorkSheets and call the .MakeExcel extension on
	/// either a List&lt;Worksheet&gt;or with a single worksheet</para>
	/// </summary>
	public class WorkBook
	{
		internal List<IXLWorksheet> worksheets { get; set; }
		/// <summary>
		/// Name of the exported file.<para>Default name is Export</para>
		/// </summary>
		public string fileName { get => _fileName; set => _fileName = value ?? "Export"; }
		private string _fileName { get; set; }
		/// <summary>
		/// Set text wrap on for the cells in the document <para>Default is true</para>
		/// </summary>
		public bool textWrap { get => _textWrap; set => _textWrap = value; }
		private bool _textWrap { get; set; }
		/// <summary>
		/// Convert a worksheet to the workbook for exporting excel
		/// </summary>
		/// <param name="ws">Worksheet to add to the workbook</param>
		/// <param name="fileName">Name of the file to be exported</param>
		/// <param name="textWrap">Wrap text <para> default is true</para></param>
		public WorkBook(WorkSheet ws, string fileName = "Export", bool textWrap = true)
		{
			worksheets = new List<IXLWorksheet>();
			_fileName = fileName;
			_textWrap = textWrap;
			convertSheets(ws);
		}
		/// <summary>
		/// Convert a list of worksheets into the workbook for exporting excel
		/// </summary>
		/// <param name="ws">List of worksheets to add to the workbook</param>
		/// <param name="fileName">Name of the file to be exported</param>
		/// <param name="textWrap">Wrap text <para> default is true</para></param>
		public WorkBook(List<WorkSheet> ws, string fileName = "Export", bool textWrap = true)
		{
			worksheets = new List<IXLWorksheet>();

			_fileName = fileName;
			_textWrap = textWrap;
			convertSheets(ws);
		}
		private void convertSheets(WorkSheet sheet)
		{
			worksheets.Add(sheet.worksheet);
		}
		private void convertSheets(List<WorkSheet> sheets)
		{
			foreach (WorkSheet sheet in sheets)
			{
				worksheets.Add(sheet.worksheet);
			}
		}
	}
}
