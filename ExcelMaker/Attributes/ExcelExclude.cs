using System;

namespace ExcelMaker.Attributes
{
	/// <summary>
	/// Add this to an attribute in a model to exclude it from the generated excel document
	/// <para>Default is true</para>
	/// </summary>
	[AttributeUsage(AttributeTargets.Property)]
	public class ExcelExcludeAttribute : Attribute
	{
		private static bool excluded1;
		/// <summary>
		/// Get the value for this property
		/// </summary>
		/// <returns></returns>
		public static bool Getexcluded()
		{
			return excluded1;
		}	

		/// <summary>
		/// Add this to an attribute in a model to exclude it from the generated excel document
		/// <para>Default is true</para>
		/// </summary>
		public ExcelExcludeAttribute(bool exclude = true)
		{
			excluded1 = exclude;
		}
	}
}
