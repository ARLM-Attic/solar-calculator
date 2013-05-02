using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System
{
	/// <summary>
	/// Defines a type for a geometric angle that allows various ways of setting and converting
	/// the values from and to different modes. Angle inherently is a double and can be used in 
	/// place of a double when specifically referring to a geometric angle with the value being
	/// the degrees of the angle.
	/// </summary>
	public class Angle : IComparable, IFormattable, IComparable<Angle>, IEquatable<Angle>
	{
		// ***
		// *** Represents the internal value of this type. This value is in degrees.
		// ***
		private double _internalValue = 0d;

		#region Constructors
		/// <summary>
		/// Initializes a new instance of the System.Angle class.
		/// </summary>
		public Angle()
		{
		}

		/// <summary>
		/// Initializes a new instance of the System.Angle class to an angle specified in degrees
		/// represented by a double.
		/// </summary>
		/// <param name="degrees">The value of the angle in degrees.</param>
		public Angle(double degrees)
		{
			this.InternalValue = degrees;
		}

		/// <summary>
		/// Initializes a new instance of the System.Angle class to an angle with the whole part
		/// specified in degrees represented by a integer and the decimal part represented in 
		/// Minutes and Seconds of an Arc. This allows an angle represented as 9° 14' 55.8'' to
		/// be initialized.
		/// </summary>
		/// <param name="degrees">The whole value portion of the angle in degrees, e.g., 9 from the angle 9° 14' 55.8''.</param>
		/// <param name="arcminute">The arcminute value, e.g., 14 from the angle 9° 14' 55.8''.</param>
		/// <param name="arcsecond">The arcsecond value, e.g., 55.8 from the angle 9° 14' 55.8''.</param>
		public Angle(int degrees, int arcminute, double arcsecond)
		{
			this.InternalValue = Angle.ToDegrees(degrees, arcminute, arcsecond);
		}
		#endregion

		#region Public Members
		/// <summary>
		/// Gets the whole number portion of the value of
		/// this instance in degrees.
		/// </summary>
		public int Degrees
		{
			get
			{
				return (int)Math.Floor(this.InternalValue);
			}
		}

		/// <summary>
		/// Gets the arcminute portion of the value of
		/// this instance in degrees.
		/// </summary>
		public int ArcmMinute
		{
			get
			{
				return Angle.GetArcminute(this);
			}
		}

		/// <summary>
		/// Gets the arcsecond portion of the value of
		/// this instance in degrees.
		/// </summary>
		public double Arcsecond
		{
			get
			{
				return Angle.GetArcsecond(this);
			}
		}

		/// <summary>
		/// Returns the value of this instance in Radian units.
		/// </summary>
		public double Radians
		{
			get
			{
				return Angle.ToRadians(this.InternalValue);
			}
		}

		/// <summary>
		/// Returns the value of this instance in Radian units by
		/// first multiplying the underlying value of the angle 
		/// (in degrees) by the multiplier before converting the
		/// value to Radian units. Note that this value is NOT
		/// equal to this.Radians * multiplier.
		/// </summary>
		public double RadiansMultiplied(double multiplier)
		{
			return Angle.ToRadians(this.InternalValue * multiplier);
		}

		/// <summary>
		/// Returns the value of this instance in Radian units by
		/// first dividing the underlying value of the angle 
		/// (in degrees) by the divisor before converting the
		/// value to Radian units. Note that this value is NOT
		/// equal to this.Radians / divisor.
		/// </summary>
		public double RadiansDivided(double divisor)
		{
			return Angle.ToRadians(this.InternalValue / divisor);
		}

		/// <summary>
		/// Gets the full value of this angle expressed only in
		/// Minutes of arc.
		/// </summary>
		public double TotalArcminutes
		{
			get
			{
				throw new NotImplementedException();
			}
		}

		/// <summary>
		/// Gets the full value of this angle expressed only in
		/// Seconds of arc.
		/// </summary>
		public double TotalArcseconds
		{
			get
			{
				throw new NotImplementedException();
			}
		}

		/// <summary>
		/// Normalizes this instance of the angle to between 0 and 360 degrees.
		/// </summary>
		public void Reduce()
		{
			this.InternalValue = Angle.Reduce(this).InternalValue;
		}

		/// <summary>
		///  Returns a value indicating whether this instance and a specified System.Angle
		/// object represent the same value.
		/// </summary>
		/// <param name="obj">A System.Angle object to compare to this instance.</param>
		/// <returns>True if obj is equal to this instance; False otherwise.</returns>
		public override bool Equals(object obj)
		{
			bool returnValue = false;

			if (obj is Angle)
			{
				Angle compare = obj as Angle;
				returnValue = this.InternalValue.Equals(compare.InternalValue);
			}

			return returnValue;
		}

		/// <summary>
		///Returns the hash code for this instance.
		/// </summary>
		/// <returns>A 32-bit signed integer hash code.</returns>
		public override int GetHashCode()
		{
			return this.InternalValue.GetHashCode();
		}
		#endregion

		#region Implicit Conversions
		/// <summary>
		/// Defines an implicit conversion of a System.Angle object to a System.Double object.
		/// </summary>
		/// <param name="angle">The object to convert.</param>
		/// <returns>The converted object.</returns>
		public static implicit operator double(Angle angle)
		{
			return angle.InternalValue;
		}

		/// <summary>
		/// Defines an implicit conversion of a System.Double object to a System.Angle object.
		/// </summary>
		/// <param name="degrees">The object to convert.</param>
		/// <returns>The converted object.</returns>
		public static implicit operator Angle(double degrees)
		{
			return new Angle(degrees);
		}
		#endregion

		#region Static Members
		/// <summary>
		/// Represents the empty Angle. This field is read-only.
		/// </summary>
		public static Angle Empty
		{
			get
			{
				return new Angle(0);
			}
		}

		/// <summary>
		/// Converts the value of an angle specified in degree, arcminute and arcsecond
		/// values to a double-precision floating-point value of the angle. For example
		/// an angle specified as as 9° 14' 55.8'' can be converted to a double.
		/// </summary>
		/// <param name="degrees">The whole value portion of the angle in degrees, e.g., 9 from the angle 9° 14' 55.8''.</param>
		/// <param name="arcminute">The arcminute value, e.g., 14 from the angle 9° 14' 55.8''.</param>
		/// <param name="arcsecond">The arcsecond value, e.g., 55.8 from the angle 9° 14' 55.8''.</param>
		/// <returns>The double-precision floating-point value of the specified angle.</returns>
		public static double ToDegrees(int degrees, int arcminute, double arcsecond)
		{
			double returnValue = 0d;

			// ***
			// *** Ensure all parameters are the same sing
			// ***
			Angle.NormalDirection(ref degrees, ref arcminute, ref arcsecond);

			// ***
			// *** 15 degrees per hour
			// ***
			returnValue = degrees + (arcminute / 60d) + (arcsecond / 3600d);

			return returnValue;
		}

		/// <summary>
		/// Converts a radian angle to degrees.
		/// </summary>
		/// <param name="radians">The angle in radian units.</param>
		/// <returns>The angle in degree units.</returns>
		public static double ToDegrees(double radians)
		{
			double returnValue = 0d;

			// ***
			// *** Factor = 180 / pi 
			// ***
			returnValue = radians * (180d / Math.PI);

			return returnValue;
		}

		/// <summary>
		/// A radian is the measure of a central angle subtending an arc equal in length to the radius. This method 
		/// converts an angle measured in degrees to radian units.
		/// </summary>
		/// <param name="degrees">The angle in degree units.</param>
		/// <returns>The angle in radian units.</returns>
		public static double ToRadians(double degrees)
		{
			double returnValue = 0d;

			// ***
			// *** Factor = pi / 180 
			// ***
			returnValue = (degrees * Math.PI) / 180d;

			return returnValue;
		}

		/// <summary>
		/// Creates an instance of the System.Angle class from the
		/// specified angle expressed in Radian units.
		/// </summary>
		/// <param name="radians">The angle in radian units.</param>
		/// <returns>A new instance of the System.Angle class.</returns>
		public static Angle FromRadians(double radians)
		{
			Angle returnValue = new Angle();

			returnValue = new Angle(Angle.ToDegrees(radians));

			return returnValue;
		}

		/// <summary>
		/// A minute of arc, arcminute, is a unit of angular measurement equal to one sixtieth (1⁄60) of one degree. This method
		/// returns the arcminute portion of the given Angle.
		/// </summary>
		/// <param name="angle">An instance of the System.Angle class.</param>
		/// <returns>The arcminute of the specified angle.</returns>
		public static int GetArcminute(Angle angle)
		{
			int returnValue = 0;

			double decimalPortion = angle - Math.Floor(angle);
			returnValue = (int)Math.Floor(decimalPortion * 60);

			return returnValue;
		}

		/// <summary>
		/// A second of arc or arcsecond is one sixtieth (1⁄60) of one arcminute. This method
		/// returns the arcsecond portion of the given Angle.
		/// </summary>
		/// <param name="angle">An instance of the System.Angle class.</param>
		/// <returns>The arcminute of the specified angle.</returns>
		public static double GetArcsecond(Angle angle)
		{
			double returnValue = 0d;

			double decimalPortion = angle - Math.Floor(angle);
			double totalMinutes = decimalPortion * 60;
			double secondsDecimal = totalMinutes - Math.Floor(totalMinutes);
			returnValue = (secondsDecimal) * 60;

			return returnValue;
		}

		/// <summary>
		/// Normalizes an angle to between 0 and 360 degrees.
		/// </summary>
		/// <param name="angle">The angle in degrees to be reduced.</param>
		/// <returns>A representation of the given Angle in degrees where the value is between 0 and 360.</returns>
		public static Angle Reduce(Angle angle)
		{
			Angle returnValue = Angle.Empty;

			returnValue = new Angle(angle.InternalValue - (Math.Floor(angle.InternalValue / 360d) * 360d));

			return returnValue;
		}

		/// <summary>
		/// Converts the string representation of a angle expressed as a number to its 
		/// double-precision floating-point number equivalent in degree units.
		/// </summary>
		/// <param name="s">A string that contains a number to convert.</param>
		/// <returns>An instance of System.Angle instantiated with the parsed value.</returns>
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
		/// Converts the string representation of a angle expressed as a number to its 
		/// double-precision floating-point number equivalent in degree units. A return 
		/// value indicates whether the conversion succeeded or failed.
		/// </summary>
		/// <param name="s">A string that contains a number to convert.</param>
		/// <param name="result">An instance of System.Angle instantiated with the parsed value.</param>
		/// <returns>True if s was converted successfully; False otherwise.</returns>
		public static bool TryParse(string s, out Angle result)
		{
			bool returnValue = false;
			result = null;

			double value = 0d;
			if (double.TryParse(s, out value))
			{
				result = new Angle(value);
				returnValue = true;
			}

			return returnValue;
		}
		#endregion

		#region Internal Members
		internal double InternalValue
		{
			get
			{
				return _internalValue;
			}
			set
			{
				_internalValue = value;
			}
		}

		internal static void NormalDirection(ref int degrees, ref int minutes, ref double seconds)
		{
			// ***
			// *** If the value for degrees is negative then
			// *** minutes and seconds should be negative. This 
			// *** is due to the fact that these three numbers
			// *** represent one angle and always must have the
			// *** same direction.
			// ***
			if (degrees < 0 || minutes < 0 || seconds < 0)
			{
				degrees = -1 * Math.Abs(degrees);
				minutes = -1 * Math.Abs(minutes);
				seconds = -1 * Math.Abs(seconds);
			}
		}

		internal static void NormalDirection(ref double hours, ref int minutes, ref double seconds)
		{
			// ***
			// *** If the value for degrees is negative then
			// *** minutes and seconds should be negative. This 
			// *** is due to the fact that these three numbers
			// *** represent one angle and always must have the
			// *** same direction.
			// ***
			if (hours < 0 || minutes < 0 || seconds < 0)
			{
				hours = -1 * Math.Abs(hours);
				minutes = -1 * Math.Abs(minutes);
				seconds = -1 * Math.Abs(seconds);
			}
		}
		#endregion

		#region Formatters
		/// <summary>
		/// Converts the value of this instance to its equivalent string representation.
		/// </summary>
		/// <returns>The string representation of the value of this instance.</returns>
		public override string ToString()
		{
			return this.ToShortFormat();
		}

		/// <summary>
		/// The string representation of the value of this instance using the common
		/// symbol for degrees and a decimal value for the portion that is less than
		/// zero.
		/// </summary>
		/// <returns>The short format string representation of the value of this instance.</returns>
		public string ToShortFormat()
		{
			return this.InternalValue.ToString("0°.0000####");
		}

		/// <summary>
		/// The string representation of the value of this instance displayed in degrees, arcminutes
		/// and arcseonds.
		/// </summary>
		/// <returns>The long format string representation of the value of this instance.</returns>
		public string ToLongFormat()
		{
			return string.Format("{0}°{1}´{2:0´´.#####}", this.Degrees, this.ArcmMinute, this.Arcsecond);
		}

		//public string ToHourFormat()
		//{
		//	return string.Format("{0}ʰ {1}ᵐ {2:0ˢ.#####}", this.Degrees, this.ArcmMinute, this.Arcsecond);
		//}
		#endregion

		#region IComparable
		/// <summary>
		/// Compares this instance with a specified System.Angle object and indicates
		/// whether this instance precedes, follows, or appears in the same position
		/// as the specified System.Angle.
		/// </summary>
		/// <param name="obj">The angle to compare with this instance.</param>
		/// <returns>A 32-bit signed integer that indicates whether this instance precedes,
		/// follows, or appears in the same position as the value parameter. Value 
		/// Condition Less than zero This instance precedes obj. Zero This instance
		/// has the same position as obj. Greater than zero This instance
		/// follows obj or obj is null.</returns>
		public int CompareTo(object obj)
		{
			int returnValue = -1;

			returnValue = this.InternalValue.CompareTo(((Angle)obj).InternalValue);

			return returnValue;
		}
		#endregion

		#region IFormattable
		/// <summary>
		/// Converts the numeric value of this instance to its equivalent string representation
		/// using the specified culture-specific format information.
		/// </summary>
		/// <param name="format">A numeric format string.</param>
		/// <param name="formatProvider">An object that supplies culture-specific formatting information.</param>
		/// <returns>The string representation of the value of this instance as specified by format and provider.</returns>
		public string ToString(string format, IFormatProvider formatProvider)
		{
			return this.InternalValue.ToString(format, formatProvider);
		}
		#endregion

		#region IComparable<Angle>
		/// <summary>
		/// Compares this instance with a specified System.Angle object and indicates
		/// whether this instance precedes, follows, or appears in the same position
		/// as the specified System.Angle.
		/// </summary>
		/// <param name="other">The angle to compare with this instance.</param>
		/// <returns>A 32-bit signed integer that indicates whether this instance precedes,
		/// follows, or appears in the same position as the value parameter. Value 
		/// Condition Less than zero This instance precedes obj. Zero This instance
		/// has the same position as obj. Greater than zero This instance
		/// follows other or other is null.</returns>
		public int CompareTo(Angle other)
		{
			return this.InternalValue.CompareTo(other.InternalValue);
		}
		#endregion

		#region IEquatable<Angle>
		/// <summary>
		/// Specifies whether this Angle and the specified Angle have the
		/// same value.
		/// </summary>
		/// <param name="other">The Angle to compare to this instance.</param>
		/// <returns>True if the value of the other parameter is the same as this Angle; False otherwise</returns>
		public bool Equals(Angle other)
		{
			return this.Equals(other);
		}
		#endregion
	}
}
