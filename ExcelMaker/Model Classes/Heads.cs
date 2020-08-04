namespace ExcelMaker.Model_Classes
{
	/// <summary>
	/// Holds the headers used in exporting excel data
	/// </summary>
	internal partial class Heads
	{
		/// <summary>
		/// Friendly display text
		/// </summary>
		internal string displayText { get; set; }
		/// <summary>
		/// Name of the field - used to get the data for the body
		/// </summary>
		internal string fieldName { get; set; }
		/// <summary>
		/// column sort order for the excel document
		/// </summary>
		internal int columnNumber { get; set; }
	}
}
