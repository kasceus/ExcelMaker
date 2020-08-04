using System;
using System.Linq;
using System.Threading.Tasks;
using ExcelMaker.Attributes;

//This class will contain the extension methods for making excel documents.
//Only the following generic classes will be extended:
// IQueryable<object>, IEnumerable<object>
namespace ExcelMaker
{
	using System.Collections.Generic;
	using System.ComponentModel.DataAnnotations;
	using System.Data;
	using System.Reflection;
	using ExcelMaker.Model_Classes;

	/// <summary>
	/// Extend the classes used in the excel maker
	/// </summary>
	public static class Extensions
	{
		#region MakeExcel Extensions
		/// <summary>
		/// Make an excel document using a worksheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="fileName"></param>
		/// <param name="textWrap"></param>
		public static void MakeExcel(this WorkSheet sheet, string fileName, bool textWrap = true)
		{
			WorkBook book = new WorkBook(sheet, fileName, textWrap);
			book.MakeExcel();
		}
		/// <summary>
		/// Make an excel document from a workbook
		/// </summary>
		/// <param name="wb"></param>
		public static void MakeExcel(this WorkBook wb)
		{
			if (wb == null)
			{
				throw new ArgumentNullException(nameof(wb));
			}

			Methods.MakeExcel(wb);
		}
		#endregion
		#region ToDataTable Extension
		/// <summary>
		/// Converts a data object into a DataTable for use with the Excel Maker
		/// </summary>
		/// <param name="data">data to format</param>
		/// <param name="tableName"></param>
		/// <returns>DataTable object for the supplied data</returns>
		public static DataTable ToDataTable(this IEnumerable<object> data, string tableName)
		{
			DataTable dt = new DataTable();
			Type type = data.GetType().GetTypeInfo().GenericTypeArguments[0];
			string fQN = type.AssemblyQualifiedName;//full assembly name required here to get the proper type castings
			List<Heads> heads = dt.MakeHeaders(fQN);
			dt.MakeRows(data, heads);
			dt.TableName = tableName;

			return dt;
		}
		#endregion
		/// <summary>
		/// data table extension used by the <see cref="Methods"/> class.  Make the columns for the datatable. Returns the ordered list of headers.
		/// </summary>
		/// <param name="table">dataTable reference</param>
		/// <param name="FQN">Fully quallified name for the object being passed in for analysis</param>
		internal static List<Heads> MakeHeaders(this DataTable table, string FQN)
		{
			Type type = Type.GetType(FQN, true);
			PropertyInfo[] props = type.GetProperties();
			List<Heads> heads = new List<Heads>();
			int counter = 1;
			//use parallelization to process all the data and make the unordered listing of the header information
			Parallel.ForEach(props, property =>
			{
				//check if exclude attribute is set for this field
				if (Attribute.GetCustomAttribute(property, typeof(ExcelExcludeAttribute)) is ExcelExcludeAttribute)
				{
					return;
				}
				string displayText = "";
				//check if display name set for this field
				if (!(Attribute.GetCustomAttribute(property, typeof(DisplayAttribute)) is DisplayAttribute display))
				{
					displayText = property.Name;
				}
				else
				{
					displayText = display.Name;
				}
				int colNumber = 0;
				if (Attribute.GetCustomAttribute(property, typeof(KeyAttribute)) is KeyAttribute)
				{
					colNumber = 0;
				}
				else
				{
					if (Attribute.GetCustomAttribute(property, typeof(ColumnNumberAttribute)) is ColumnNumberAttribute)
					{
						ColumnNumberAttribute colAttr = Attribute.GetCustomAttribute(property, typeof(ColumnNumberAttribute)) as ColumnNumberAttribute;
						colNumber = colAttr.columnNumber;
					}
					else
					{
						colNumber = counter;
					}
				}
				//check if column number attribute set for this field

				//add the field data to the head listing
				Heads head = new Heads()
				{
					columnNumber = colNumber,
					displayText = displayText,
					fieldName = property.Name
				};
				lock (heads)
				{
					heads.Add(head);
				}
				counter++;
			});
			//Order the data that was processed out of order (inherent problem with parallelization)
			heads = heads.OrderBy(a => a.columnNumber).ToList();
			//add the ordered table data to the dataTable
			foreach (Heads head in heads)
			{
				table.Columns.Add(head.displayText);
			}
			return heads;
		}
		/// <summary>
		/// Make the rows for the exported excel document
		/// </summary>
		/// <param name="table"></param>
		/// <param name="dataSet"></param>
		/// <param name="heads"></param>
		internal static void MakeRows(this DataTable table, IEnumerable<object> dataSet, List<Heads> heads)
		{
			List<Row> rows = new List<Row>();
			//loop through each row and create the rows and row data in parallel
			int colStart = 10;
			Parallel.ForEach(dataSet, data =>
			{
				IList<PropertyInfo> props = data.GetType().GetProperties();
				Row row = new Row();
				List<RowData> rowDatas = new List<RowData>();

				foreach (PropertyInfo prop in props)
				{
					if (Attribute.GetCustomAttribute(prop, typeof(ExcelExcludeAttribute)) is ExcelExcludeAttribute)
					{
						continue;
					}
					if (Attribute.GetCustomAttribute(prop, typeof(KeyAttribute)) as KeyAttribute != null)
					{
						row.recordId = prop.GetValue(data).ToString();
					}
					foreach (Heads a in heads)
					{
						if (a.fieldName == prop.Name)
						{
							RowData rowData = new RowData
							{
								columnNumber = a.columnNumber
							};
							if (rowData.columnNumber <colStart)
							{
								colStart = rowData.columnNumber;
							}
							object val = prop.GetValue(data, null);
							rowData.data = (val != null)
							? (val.GetType() == typeof(DateTime))
							? ((DateTime)val).ToShortDateString()
							: val.ToString()
							: "";
							rowDatas.Add(rowData);
						}
					}
				}
				row.data = rowDatas;
				lock (rows)
				{
					rows.Add(row);
				}
			});

			//loop through the rows and add them to the DataTable object in parallel
			foreach (Row row in rows)
			{
				DataRow tRow = table.NewRow();
				IOrderedEnumerable<RowData> colSorted = row.data.OrderBy(b => b.columnNumber);
				foreach (RowData d in colSorted)
				{//can't parallelize this since column order matters
					if (colStart == 0)
					{
						tRow[d.columnNumber] = d.data ?? "";
					}
					else
					{
						tRow[d.columnNumber-colStart] = d.data ?? "";
					}
				}
				table.Rows.Add(tRow);
			}

		}
	}
}
