using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System
{
	/// <summary>
	/// Defines a type for a geometric angle that allows various ways of setting and converting
	/// the values from and to different modes. Angle inherently is a double and can be used in 
	/// place of a double when specifically referring to a geometric angle.
	/// </summary>
	public class Angle : IComparable, IFormattable, IComparable<Angle>, IEquatable<Angle>
	{
		// ***
		// *** This value is stored in degrees
		// ***
		private double _value = 0.0;

		/// <summary>
		/// 
		/// </summary>
		public Angle()
		{
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		public Angle(double degrees)
		{
			this.Value = degrees;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		/// <param name="minutes"></param>
		/// <param name="seconds"></param>
		public Angle(int degrees, int minutes, double seconds)
		{
			this.Value = Angle.ToDegrees(degrees, minutes, seconds);
		}

		internal double Value
		{
			get
			{
				return _value;
			}
			set
			{
				_value = value;
			}
		}

		#region Public Members
		/// <summary>
		/// 
		/// </summary>
		public int Degrees
		{
			get
			{
				return (int)Math.Floor(this.Value);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public int Minutes
		{
			get
			{
				return Angle.GetMinutes(Angle.ToHours(this.Value));
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public double Seconds
		{
			get
			{
				return Angle.GetSeconds(Angle.ToHours(this.Value));
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public double Radians
		{
			get
			{
				return Angle.ToRadians(this.Value);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public double Hours
		{
			get
			{
				return Angle.ToHours(this.Value);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public override bool Equals(object obj)
		{
			bool returnValue = false;

			if (obj is Angle)
			{
				Angle compare = obj as Angle;
				returnValue = this.Value.Equals(compare.Value);
			}

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public override int GetHashCode()
		{
			return this.Value.GetHashCode();
		}
		#endregion

		#region Implicit Conversions
		/// <summary>
		/// 
		/// </summary>
		/// <param name="angle"></param>
		/// <returns></returns>
		public static implicit operator double(Angle angle)
		{
			return angle.Value;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		/// <returns></returns>
		public static implicit operator Angle(double degrees)
		{
			return new Angle(degrees);
		}
		#endregion

		#region Static Members
		/// <summary>
		/// 
		/// </summary>
		public static Angle Empty
		{
			get
			{
				return new Angle(0);
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		/// <param name="minutes"></param>
		/// <param name="seconds"></param>
		/// <returns></returns>
		public static double ToHours(int degrees, int minutes, double seconds)
		{
			double returnValue = 0.0;

			returnValue = (degrees / 15.0) + (minutes / 60.0) + (seconds / 3600.0);

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="hours"></param>
		/// <param name="minutes"></param>
		/// <param name="seconds"></param>
		/// <returns></returns>
		public static double ToTotalHours(int hours, int minutes, double seconds)
		{
			double returnValue = 0.0;

			returnValue = hours + (minutes / 60.0) + (seconds / 3600.0);

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		/// <returns></returns>
		public static double ToHours(double degrees)
		{
			return degrees / 15.0;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="degrees"></param>
		/// <param name="minutes"></param>
		/// <param name="seconds"></param>
		/// <returns></returns>
		public static double ToDegrees(int degrees, int minutes, double seconds)
		{
			double returnValue = 0.0;

			// ***
			// *** If the value for degrees is negative then
			// *** minutes and seconds should be negative. This 
			// *** is due to the fact that these three numbers
			// *** represent one angle and always must have the
			// *** same direction.
			// ***
			if (degrees < 0)
			{
				minutes = -1 * Math.Abs(minutes);
				seconds = -1 * Math.Abs(seconds);
			}

			// ***
			// *** 15 degrees per hour
			// ***
			returnValue = Angle.ToHours(degrees, minutes, seconds) * 15.0;

			return returnValue;
		}

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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="radians"></param>
		/// <returns></returns>
		public static Angle FromRadians(double radians)
		{
			Angle returnValue = new Angle();

			returnValue = new Angle(Angle.ToDegrees(radians));

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="hours"></param>
		/// <returns></returns>
		public static Angle FromHours(double hours)
		{
			Angle returnValue = null;

			returnValue = hours * 15.0;

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="hours"></param>
		/// <returns></returns>
		public static int GetDegrees(double hours)
		{
			return (int)Math.Floor(hours);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="hours"></param>
		/// <returns></returns>
		public static int GetMinutes(double hours)
		{
			return (int)Math.Floor((hours - Angle.GetDegrees(hours)) * 60.0);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="hours"></param>
		/// <returns></returns>
		public static double GetSeconds(double hours)
		{
			double returnValue = 0.0;

			int degrees = Angle.GetDegrees(hours);
			int minutes = Angle.GetMinutes(hours);
			returnValue = (hours - degrees - (minutes / 60.0)) * 3600.0;

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		public static Angle Parse(string s)
		{
			Angle returnValue = Angle.Empty;

			if (!Angle.TryParse(s, out returnValue))
			{
				throw new FormatException();
			}

			return returnValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="s"></param>
		/// <param name="result"></param>
		/// <returns></returns>
		public static bool TryParse(string s, out Angle result)
		{
			bool returnValue = false;
			result = null;

			return returnValue;
		}
		#endregion

		#region Formatters
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public override string ToString()
		{
			return this.ToShortFormat();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public string ToShortFormat()
		{
			return string.Format("{0}", this.Value.ToString("0°.0000"));
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public string ToLongFormat()
		{
			return string.Format("{0}° {1}' {2}\"", this.Degrees, this.Minutes, this.Seconds.ToString("0.0"));
		}

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public string ToHourFormat()
		{
			return string.Format("{0}h {1}m {2}s", Math.Floor(this.Hours), this.Minutes, this.Seconds.ToString("0.0"));
		}
		#endregion

		#region IComparable
		/// <summary>
		/// 
		/// </summary>
		/// <param name="obj"></param>
		/// <returns></returns>
		public int CompareTo(object obj)
		{
			int returnValue = -1;

			if (obj is Angle)
			{
				returnValue = this.Value.CompareTo(((Angle)obj).Value);
			}

			return returnValue;
		}
		#endregion

		#region IFormattable
		/// <summary>
		/// 
		/// </summary>
		/// <param name="format"></param>
		/// <param name="formatProvider"></param>
		/// <returns></returns>
		public string ToString(string format, IFormatProvider formatProvider)
		{
			return this.Value.ToString(format, formatProvider);
		}
		#endregion

		#region IComparable<Time>
		/// <summary>
		/// 
		/// </summary>
		/// <param name="other"></param>
		/// <returns></returns>
		public int CompareTo(Angle other)
		{
			return this.Value.CompareTo(other.Value);
		}
		#endregion

		#region IEquatable<Time>
		/// <summary>
		/// 
		/// </summary>
		/// <param name="other"></param>
		/// <returns></returns>
		public bool Equals(Angle other)
		{
			return this.Equals(other);
		}
		#endregion
	}
}
