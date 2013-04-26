using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System
{
	/// <summary>
	/// Provides methods to convert angles between Degrees and Radians.
	/// </summary>
	public class Angle
	{
		/// <summary>
		/// Converts a radian angle to degrees
		/// </summary>
		/// <param name="radians">The angle in radian units</param>
		/// <returns>The angle in degree units</returns>
		public static double ToDegrees(double radians)
		{
			double returnValue = 0.0;

			// ***
			// *** Factor = 180 / pi 
			// ***
			returnValue = radians * (180.0 / Math.PI);

			return returnValue;
		}

		/// <summary>
		/// A radian is the measure of a central angle subtending an arc equal in length to the radius. This method 
		/// converts an angle measured in degrees to radian units.
		/// </summary>
		/// <param name="degrees">The angle in degree units</param>
		/// <returns>The angle in radian units</returns>
		public static double ToRadians(double degrees)
		{
			double returnValue = 0.0;

			// ***
			// *** Factor = pi / 180 
			// ***
			returnValue = (degrees * Math.PI) / 180.0;

			return returnValue;
		}
	}
}
