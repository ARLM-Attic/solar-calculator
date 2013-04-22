﻿// ***
// *** Copyright (C) 2013, Daniel M. Porrey.  All rights reserved.
// *** Written By Daniel M. Porrey
// ***
// *** This software is provided "AS IS," without a warranty of any kind. ALL EXPRESS OR IMPLIED CONDITIONS, REPRESENTATIONS AND WARRANTIES, 
// *** INCLUDING ANY IMPLIED WARRANTY OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR NON-INFRINGEMENT, ARE HEREBY EXCLUDED. DANIEL M PORREY 
// *** AND ITS LICENSORS SHALL NOT BE LIABLE FOR ANY DAMAGES SUFFERED BY LICENSEE AS A RESULT OF USING, MODIFYING OR DISTRIBUTING THIS SOFTWARE 
// *** OR ITS DERIVATIVES. IN NO EVENT WILL DANIEL M PORREY OR ITS LICENSORS BE LIABLE FOR ANY LOST REVENUE, PROFIT OR DATA, OR FOR DIRECT, INDIRECT, 
// *** SPECIAL, CONSEQUENTIAL, INCIDENTAL OR PUNITIVE DAMAGES, HOWEVER CAUSED AND REGARDLESS OF THE THEORY OF LIABILITY, ARISING OUT OF THE USE OF 
// *** OR INABILITY TO USE THIS SOFTWARE, EVEN IF DANIEL M PORREY HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 
// ***
// *** Licensed under Microsoft Reciprocal License (Ms-RL)
// *** This license governs use of the accompanying software. If you use the software, you accept this license. If you do not accept the license, 
// *** do not use the software. Full license details can be found at https://solarcalculator.codeplex.com/license.
// ***
namespace System
{
    /// <summary>
    /// Provides extensions used to convert datetime values used in Excel and Ole Automation.
    /// </summary>
	public static class DateTimeExtensions
	{
		/// <summary>
		/// An OLE Automation date is implemented as a floating-point number whose integral component is the number of 
		/// days before or after midnight, 30 December 1899, and whose fractional component represents the time on that 
		/// day divided by 24. For example, midnight, 31 December 1899 is represented by 1.0; 6 A.M., 1 January 1900 is 
		/// represented by 2.25; midnight, 29 December 1899 is represented by -1.0; and 6 A.M., 29 December 1899 is 
		/// represented by -1.25. The base OLE Automation Date is midnight, 30 December 1899. The minimum OLE Automation 
		/// date is midnight, 1 January 0100. The maximum OLE Automation Date is the same as DateTime.MaxValue, the last
		/// moment of 31 December 9999. The ToOADate method throws an OverflowException if the current instance represents 
		/// a date that is later than MinValue and earlier than midnight on January1, 0100. However, if the value of the 
		/// current instance is MinValue, the method returns 0.
		/// </summary>
		/// <param name="value">A DateTime value that will be converted to the Ole Automation date.</param>
		/// <returns>Gets a value that represents the Ole Automation date for the given DateTime value.</returns>
		public static double ToOleAutomationDate(this DateTime value)
		{
			return value.Subtract(new DateTime(1899, 12, 30).Date).TotalDays;
		}

		/// <summary>
		/// The DATEVALUE function converts a date that is stored as text to a serial number that Excel recognizes as a date. 
		/// For example, the formula = DATEVALUE("1/1/2008") returns 39448, the serial number of the date 1/1/2008.
		/// Note The serial number returned by the DATEVALUE function can vary from the preceding example, depending 
		/// on your computer's system date settings.
		/// The DATEVALUE function is helpful in cases where a worksheet contains dates in a text format that you want 
		/// to filter, sort, or format as dates, or use in date calculations.
		/// </summary>
		/// <param name="value">A DateTime value that will be converted to the DateValue.</param>
		/// <returns>Gets a value that represents the Excel DateValue for the given DateTime value.</returns>
		public static double ToExcelDateValue(this DateTime value)
		{
			double returnValue = 0;

			if (value.Date <= DateTime.Parse("1/1/1900"))
			{
				double d = value.ToOleAutomationDate();
				double c = Math.Floor(d);
				returnValue = 1 + (d - c);
			}
			else
			{
				returnValue = value.ToOleAutomationDate();
			}

			return returnValue;
		}
	}
}
