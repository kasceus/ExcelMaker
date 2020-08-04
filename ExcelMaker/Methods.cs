using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Web;
using ExcelMaker.Model_Classes;

namespace ExcelMaker
{
	/// <summary>
	/// Methods used for making excel documents
	/// <para>For best results, use the attributes included in the excel maker.</para>
	/// <!--Impliments IDisposable to ensure the DataTable and XLWorkbook objects get disposed of properly at GC-->
	/// </summary>
	public class Methods
	{
		internal static void MakeExcel(WorkBook workbook)
		{
			XLWorkbook wb = new XLWorkbook();
			foreach (IXLWorksheet sheet in workbook.worksheets)
			{
				wb.AddWorksheet(sheet);
			}
			ProcessSheets(workbook.textWrap, wb);
			SendWorkbook(workbook.fileName, wb);
		}
		/// <summary>
		/// Make a single worksheet from a list of DataTables for use in an excel document
		/// </summary>
		/// <param name="tables"></param>
		/// <returns></returns>
		internal static IXLWorksheet MakeWorksheet(List<DataTable> tables)
		{
			XLWorkbook wbd = new XLWorkbook();
			IXLWorksheet worksheet = wbd.AddWorksheet("Sheet");
			int currentRow = 1;
			foreach (DataTable table in tables)
			{
				IXLCell titleCell = worksheet.Cell(currentRow, 1);
				titleCell.Value = table.TableName??"";
				titleCell.Style.Font.FontSize = 20;
				titleCell.Style.Font.SetBold();
				currentRow++;
				worksheet.Cell(currentRow, 1).InsertTable(table);
				currentRow += table.Rows.Count + 3;
			}
			return worksheet;
		}
		/// <summary>
		/// Create a single worksheet for a single data table
		/// </summary>
		/// <param name="table"></param>
		/// <returns></returns>
		internal static IXLWorksheet MakeWorksheet(DataTable table)
		{
			XLWorkbook wbd = new XLWorkbook();
			IXLWorksheet worksheet = wbd.AddWorksheet("Sheet");
			int currentRow = 1;
			IXLCell titleCell = worksheet.Cell(currentRow, 1);
			titleCell.Value = table.TableName??"";
			titleCell.Style.Font.FontSize = 20;
			titleCell.Style.Font.SetBold();
			currentRow++;
			worksheet.Cell(currentRow, 1).InsertTable(table);
			return worksheet;
		}

		/// <summary>
		/// handle the formatting of the cells dependent upon the data type contained within
		/// </summary>
		private static void ProcessSheets(bool wrap, XLWorkbook wb)
		{
			IXLWorksheets worksheets = wb.Worksheets;
			List<string> sheetNames = new List<string>();
			int counter = 1;
			foreach (IXLWorksheet ws in worksheets)
			{
				if (sheetNames.Contains(ws.Name))
				{
					ws.Name += $"_{counter}";
					counter++;
				}
				sheetNames.Add(ws.Name);

				ws.Columns().AdjustToContents();//adjust widths to fit contents
				foreach (IXLColumn column in ws.Columns())
				{
					column.Style.Alignment.WrapText = wrap;
					if (column.Width > 50)
					{
						column.Width = 50;
						if (!wrap)
						{
							column.Style.Alignment.WrapText = true;
						}
					}

				}
				IXLRow firstRowUsed = ws.FirstRowUsed();//get the first row used
				IXLRangeRow currentRow = firstRowUsed.RowUsed();//this is the header column
				currentRow = currentRow.RowBelow();//go the the first row below the header row and start processing
				int lastRow = ws.LastRowUsed().RowUsed().RowNumber();
				int lastCol = ws.LastColumnUsed().ColumnNumber();
				while (currentRow.RowNumber() <= lastRow)
				{
					if (currentRow.IsEmpty())
					{
						currentRow = currentRow.RowBelow();
						continue;
					}
					Parallel.For(1, lastCol, i =>
					{
						IXLCell cell = currentRow.Cell(i);
						if (cell.Value == null)
						{
							return;//no need to process this cell
						}
						XLDataType celltype = CheckDataType(cell.Value.ToString());
						cell.DataType = celltype;
						switch (celltype)
						{
							case XLDataType.Number:

								if (cell.Value.ToString().Contains("."))
								{
									cell.Style.NumberFormat.NumberFormatId = 2;
								}
								else
								{
									cell.Style.NumberFormat.NumberFormatId = 1;
								}
								break;
							
						}
					});
					currentRow = currentRow.RowBelow();//goto the next row
				}
			}
			//Parallel.ForEach(worksheets, ws =>
			//{//process each sheet simultaniously
			//	lock (sheetNames)
			//	{
			//		if (sheetNames.Contains(ws.Name))
			//		{
			//			ws.Name += $"_{counter}";
			//			counter++;
			//		}
			//		sheetNames.Add(ws.Name);
			//	}
			//	ws.Columns().AdjustToContents();//adjust widths to fit contents
			//	foreach (IXLColumn column in ws.Columns())
			//	{
			//		column.Style.Alignment.WrapText = wrap;
			//		if (column.Width > 50)
			//		{
			//			column.Width = 50;
			//			if (!wrap)
			//			{
			//				column.Style.Alignment.WrapText = true;
			//			}
			//		}

			//	}
			//	IXLRow firstRowUsed = ws.FirstRowUsed();//get the first row used
			//	IXLRangeRow currentRow = firstRowUsed.RowUsed();//this is the header column
			//	currentRow = currentRow.RowBelow();//go the the first row below the header row and start processing
			//	int lastRow = ws.LastRowUsed().RowUsed().RowNumber();
			//	int lastCol = ws.LastColumnUsed().ColumnNumber();
			//	while (currentRow.RowNumber() <= lastRow)
			//	{
			//		if (currentRow.IsEmpty())
			//		{
			//			currentRow = currentRow.RowBelow();
			//			continue;
			//		}
			//		Parallel.For(1, lastCol, i =>
			//		{
			//			IXLCell cell = currentRow.Cell(i);
			//			if (cell.Value == null)
			//			{
			//				return;//no need to process this cell
			//			}
			//			XLDataType celltype = CheckDataType(cell.Value.ToString());
			//			cell.DataType = celltype;
			//			switch (celltype)
			//			{
			//				case XLDataType.Number:

			//					if (cell.Value.ToString().Contains("."))
			//					{
			//						cell.Style.NumberFormat.NumberFormatId = 2;
			//					}
			//					else
			//					{
			//						cell.Style.NumberFormat.NumberFormatId = 1;
			//					}
			//					break;
			//				case XLDataType.DateTime:
			//					cell.Style.NumberFormat.Format = "mm/dd/yyyy";
			//					break;
			//			}
			//		});
			//		currentRow = currentRow.RowBelow();//goto the next row
			//	}
			//});

			sheetNames.Clear();
		}
		/// <summary>
		/// Get the data type for the supplied cell data
		/// </summary>
		/// <param name="cellData">Cell data to examine</param>
		/// <returns></returns>
		private static XLDataType CheckDataType(string cellData)
		{
			if (int.TryParse(cellData, out int i))
			{
				return XLDataType.Number;
			}
			if (bool.TryParse(cellData, out bool j))
			{
				return XLDataType.Boolean;
			}
			if (DateTime.TryParse(cellData, out DateTime dt))
			{
				return XLDataType.DateTime;
			}
			return XLDataType.Text;
		}
		/// <summary>
		/// Saves the workbook to the memorystream and sends to the HttpResponse.OutputStream
		/// </summary>
		/// <param name="fileName">Name of the file</param>
		/// <param name="wb">Workbook to send to the response</param>
		private static void SendWorkbook(string fileName, XLWorkbook wb)
		{
			HttpResponse Response = HttpContext.Current.Response;
			//set headers for the packet so the browser knows what to expect
			Response.Clear();
			Response.Buffer = true;
			Response.Charset = "";
			Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			Response.AddHeader("content-disposition", $"attachment; filename={fileName}.xlsx");
			Response.ContentType = "application/vnd.ms-excel";
			MemoryStream myMemoryStream = new MemoryStream();
			try
			{
				wb.SaveAs(myMemoryStream, new SaveOptions()
				{
					ValidatePackage = false
				});
				myMemoryStream.WriteTo(Response.OutputStream);
				Response.Flush();
				Response.End();
			}			
			finally
			{
				myMemoryStream.Dispose();
			}
		}
	}
}
