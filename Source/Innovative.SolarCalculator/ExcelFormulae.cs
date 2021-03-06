﻿using System;

namespace Innovative.SolarCalculator
{
	/// <summary>
	/// Provides methods that mimic functions found in Excel the are either not available
	/// in the .NET framework or have different behavior.
	/// </summary>
	public class ExcelFormulae
	{
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
		public static decimal ToExcelDateValue(DateTime value)
		{
			decimal returnValue = 0;

			if (value.Date <= DateTime.Parse("1/1/1900"))
			{
#if NET20
				decimal d = DateTimeExtensions.ToOleAutomationDate(value);
#else
				decimal d = value.ToOleAutomationDate();
#endif
				decimal c = (decimal)Math.Floor((double)d);
				returnValue = 1M + (d - c);
			}
			else
			{
#if NET20
				returnValue = DateTimeExtensions.ToOleAutomationDate(value);
#else
				returnValue = value.ToOleAutomationDate();
#endif
			}

			return returnValue;
		}

		/// <summary>
		/// Gets returns the remainder after a number is divided by a divisor. In Microsoft Excel, the result returned 
		/// by the worksheet MOD function may be different from the result returned by the c# Mod operator. This problem 
		/// occurs if you use the MOD function with either a negative number or a negative divisor, but not both negative.
		/// See http://support.microsoft.com/kb/141178?wa=wsignin1.0 for more information.
		/// </summary>
		/// <param name="number">The numeric value whose remainder you wish to find.</param>
		/// <param name="divisor">The number used to divide the number parameter. If the divisor is 0</param>
		/// <returns></returns>
		public static decimal Mod(decimal number, decimal divisor)
		{
			decimal returnValue = 0M;

			if (divisor != 0M)
			{
				returnValue = number - divisor * (decimal)Math.Floor((double)(number / divisor));
			}
			else
			{
				throw new DivideByZeroException("The value for divisor cannot be zero.");
			}

			return returnValue;
		}
	}
}
