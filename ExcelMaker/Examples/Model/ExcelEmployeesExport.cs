using ExcelMaker.Attributes;
using System.ComponentModel.DataAnnotations;

namespace ExcelMaker.Examples.Model
{
	public partial class ExcelEmployeesExport
	{
		public ExcelEmployeesExport() { }
		[ColumnNumber(1)]
		[Display(Name = "First Name")]
		public string firstName { get; set; }
		[ColumnNumber(0)]
		[Display(Name = "Last Name")]
		public string lastName { get; set; }
		[ColumnNumber(2)]
		[Display(Name = "Rank")]
		public string rank { get; set; }
		[ColumnNumber(3)]
		[Display(Name = "Division")]
		public string division { get; set; }
		[ColumnNumber(4)]
		[Display(Name = "Section")]
		public string sectionName { get; set; }
		[ColumnNumber(5)]
		[Display(Name = "Status")]
		public string status { get; set; }
		[ColumnNumber(6)]
		[Display(Name = "Worker Type")]
		public string workerType { get; set; }
		[ColumnNumber(7)]
		[Display(Name = "From Date")]
		public string fromDate { get; set; }
		[ColumnNumber(8)]
		[Display(Name = "To Date")]
		public string toDate { get; set; }
		[ColumnNumber(9)]
		[Display(Name = "Notes")]
		public string notes { get; set; }
	}
}