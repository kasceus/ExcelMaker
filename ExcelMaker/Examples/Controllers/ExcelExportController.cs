using ExcelMaker.Examples.Model;
using ExcelMaker.Model_Classes;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.Mvc;

namespace ExcelMaker.Examples.Controllers
{
	public class ExcelExportController : Controller
	{
		//private readonly DbName db = new DbName();

		[ValidateAntiForgeryToken]
		public void exportExcel_EmployeeData()
		{
			////Get the data from the database that has a name of employees
			//List<ExcelEmployeesExport> data = (from s in db.Employees
			//								   select new ExcelEmployeesExport
			//								   {
			//									   firstName = s.FirstName,
			//									   lastName = s.LastName,
			//									   rank = s.RankGrade.RankGrade1,
			//									   division = s.Section.Division.Division1,
			//									   sectionName = s.Section.SectionName,
			//									   status = s.DutyStatu.Status,
			//									   workerType = s.WorkerType.WorkerType1,
			//									   fromDate = s.FromDate.ToString() ?? "",
			//									   toDate = s.ToDate.ToString() ?? "",
			//									   notes = s.Notes
			//								   }).OrderBy(a => a.lastName).ToList();
			//List<WorkSheet> sheets = new List<WorkSheet>
			//{
			//	new WorkSheet(data, "All Employees")
			//};
			////creating the worksheets by division
			//foreach (IGrouping<string, ExcelEmployeesExport> b in data.GroupBy(a => a.division))
			//{
			//	List<DataTable> divisionTables = new List<DataTable>();
			//	IOrderedEnumerable<IGrouping<string, ExcelEmployeesExport>> sections = b.GroupBy(a => a.sectionName).OrderBy(d => d.Key);
			//	foreach (IGrouping<string, ExcelEmployeesExport> section in sections)
			//	{
			//		DataTable dt = data
			//			.Where(a => a.sectionName == section.Key)
			//			.OrderBy(a => a.lastName)
			//			.ToList()
			//			.ToDataTable(section.Key);
			//		divisionTables.Add(dt);
			//	}
			//	sheets.Add(new WorkSheet(divisionTables, b.Key));
			//}
			//WorkBook wb = new WorkBook(sheets, "Employee Listing");
			//wb.MakeExcel();
		}
	}
}