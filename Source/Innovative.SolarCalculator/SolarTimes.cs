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
using System;

namespace Innovative.SolarCalculator
{
	/// <summary>
	/// Provides mathematical operations to calculate the sunrise and sunset for a given date.
	/// </summary>
	public class SolarTimes
	{
		private DateTimeOffset _forDate = DateTimeOffset.MinValue;
		private double _atmosphericRefraction = 0.833;
		private double _longitude = 0;
		private double _latitude = 0;

		/// <summary>
		/// Creates a default instance of the SolarTimes object.
		/// </summary>
		public SolarTimes()
		{
			this.ForDate = DateTime.Now;
		}

		/// <summary>
		/// Creates an instance of the SolarTimes object with the specified ForDate.
		/// </summary>
		/// <param name="forDate">Specifies the Date for which the sunrise and sunset will be calculated.</param>
		public SolarTimes(DateTimeOffset forDate)
		{
			this.ForDate = forDate;
		}

		/// <summary>
		/// Specifies the Date for which the sunrise and sunset will be calculated.
		/// </summary>
		public DateTimeOffset ForDate
		{
			get
			{
				return _forDate;
			}
			set
			{
				_forDate = value;
			}
		}

		/// <summary>
		/// Angular measurement of east-west location on Earth's surface. Longitude is defined from the 
		/// prime meridian, which passes through Greenwich, England. The international date line is defined 
		/// around +/- 180° longitude. (180° east longitude is the same as 180° west longitude, because 
		/// there are 360° in a circle.) Many astronomers define east longitude as positive. This 
		/// solar calculator conforms to the international standard, with east longitude positive.
		/// (Spreadsheet Column B, Row 4)
		/// </summary>
		public double Longitude
		{
			get
			{
				return _longitude;
			}
			set
			{
				if (value >= -180 && value <= 180)
				{
					_longitude = value;
				}
				else
				{
					throw new ArgumentOutOfRangeException("The value for Longitude must be between -180° and 180°.");
				}
			}
		}

		/// <summary>
		/// Angular measurement of north-south location on Earth's surface. Latitude ranges from 90° 
		/// south (at the south pole; specified by a negative angle), through 0° (all along the equator), 
		/// to 90° north (at the north pole). Latitude is usually defined as a positive value in the 
		/// northern hemisphere and a negative value in the southern hemisphere.
		/// (Spreadsheet Column B, Row 3)
		/// </summary>
		public double Latitude
		{
			get
			{
				return _latitude;
			}
			set
			{
				if (value >= -90 && value <= 90)
				{
					_latitude = value;
				}
				else
				{
					throw new ArgumentOutOfRangeException("The value for Latitude must be between -90° and 90°.");
				}
			}
		}

		/// <summary>
		/// Gets the time zone offset for the specified date.
		/// Time Zones are longitudinally defined regions on the Earth that keep a common time. A time 
		/// zone generally spans 15° of longitude, and is defined by its offset (in hours) from UTC. 
		/// For example, Mountain Standard Time (MST) in the US is 7 hours behind UTC (MST = UTC - 7).
		/// (Spreadsheet Column B, Row 5)
		/// </summary>
		public double TimeZoneOffset
		{
			get
			{
				return this.ForDate.Offset.TotalHours;
			}
		}

		/// <summary>
		/// Sun Rise Time  
		/// (Spreadsheet Column Y)
		/// </summary>
		public DateTime Sunrise
		{
			get
			{
				DateTime returnValue = DateTime.MinValue;

				double dayFraction = this.SolarNoon.TimeOfDay.TotalDays - this.HourAngleSunrise * 4 / 1440;
				returnValue = this.ForDate.Date.Add(TimeSpan.FromDays(dayFraction));

				return returnValue;
			}
		}

		/// <summary>
		/// Sun Set Time
		/// (Spreadsheet Column Z)
		/// </summary>
		public DateTime Sunset
		{
			get
			{
				DateTime returnValue = DateTime.MinValue;

				double dayFraction = this.SolarNoon.TimeOfDay.TotalDays + this.HourAngleSunrise * 4 / 1440;
				returnValue = this.ForDate.Date.Add(TimeSpan.FromDays(dayFraction));

				return returnValue;
			}
		}

		#region Computational Members
		/// <summary>
		/// Time past local midnight.
		/// (Spreadsheet Column E)
		/// </summary>	
		public double TimePastLocalMidnight
		{
			get
			{
				double returnValue = 0.0;

				// ***
				// *** .1 / 24
				// ***
				returnValue = DateTime.Parse("12/30/1899  12:00:00 AM").Add(this.ForDate.TimeOfDay).ToOleAutomationDate();

				return returnValue;
			}
		}

		/// <summary>
		/// Julian Day: a time period used in astronomical circles, defined as the number of days 
		/// since 1 January, 4713 BCE (Before Common Era), with the first day defined as Julian 
		/// day zero. The Julian day begins at noon UTC. Some scientists use the term Julian day 
		/// to mean the numerical day of the current year, where January 1 is defined as day 001. 
		/// (Spreadsheet Column F)
		/// </summary>
		public double JulianDay
		{
			get
			{
				double returnValue = 0.0;

				// ***
				// *** this.TimePastLocalMidnight was removed since the time is in ForDate
				// ***
				returnValue = ExcelFormulae.ToExcelDateValue(this.ForDate.Date) + 2415018.5 - (this.TimeZoneOffset / 24);

				return returnValue;
			}
		}

		/// <summary>
		/// Julian Century
		/// calendar established by Julius Caesar in 46 BC, setting the number of days in a year at 365, 
		/// except for leap years which have 366, and occurred every 4 years. This calendar was reformed 
		/// by Pope Gregory XIII into the Gregorian calendar, which further refined leap years and corrected 
		/// for past errors by skipping 10 days in October of 1582. 
		/// (Spreadsheet Column G)
		/// </summary>
		public double JulianCentury
		{
			get
			{
				double returnValue = 0.0;

				returnValue = (this.JulianDay - 2451545) / 36525;

				return returnValue;
			}
		}

		/// <summary>
		/// Sun's Geometric Mean Longitude (degrees): Geometric Mean Ecliptic Longitude of Sun.
		/// (Spreadsheet Column I)
		/// </summary>
		public double SunGeometricMeanLongitude
		{
			get
			{
				double returnValue = 0.0;

				returnValue = ExcelFormulae.Mod(280.46646 + this.JulianCentury * (36000.76983 + this.JulianCentury * 0.0003032), 360);

				return returnValue;
			}
		}

		/// <summary>
		/// Sun's Mean Anomaly (degrees): Position of Sun relative to perigee
		/// (Spreadsheet Column J)
		/// </summary>
		public double SunMeanAnomaly
		{
			get
			{
				double returnValue = 0.0;

				returnValue = 357.52911 + this.JulianCentury * (35999.05029 - 0.0001537 * this.JulianCentury);

				return returnValue;
			}
		}

		/// <summary>
		/// Eccentricity of the Earth's Orbit: Eccentricity e is the ratio of half the distance between the foci c to
		/// the semi-major axis a: e = c / a. For example, an orbit with e = 0 is circular, e = 1 is parabolic, and e 
		/// between 0 and 1 is elliptic.
		/// (Spreadsheet Column K)
		/// </summary>
		public double EccentricityOfEarthOrbit
		{
			get
			{
				double returnValue = 0.0;

				returnValue = 0.016708634 - this.JulianCentury * (0.000042037 + 0.0000001267 * this.JulianCentury);

				return returnValue;
			}
		}

		/// <summary>
		/// Sun Equation of the Center: Difference between mean anomaly and true anomaly.
		/// (Spreadsheet Column L)
		/// </summary>
		public double SunEquationOfCenter
		{
			get
			{
				double returnValue = 0.0;

				returnValue = Math.Sin(Angle.ToRadians(SunMeanAnomaly)) * (1.914602 - this.JulianCentury * (0.004817 + 0.000014 * this.JulianCentury)) + Math.Sin(Angle.ToRadians(2.0 * SunMeanAnomaly)) * (0.019993 - 0.000101 * JulianCentury) + Math.Sin(Angle.ToRadians(3.0 * this.SunMeanAnomaly)) * 0.000289;

				return returnValue;
			}
		}

		/// <summary>
		/// Sun True Long (degrees)
		/// (Spreadsheet Column M)
		/// </summary>
		public double SunTrueLongitude
		{
			get
			{
				double returnValue = 0.0;

				returnValue = this.SunGeometricMeanLongitude + this.SunEquationOfCenter;

				return returnValue;
			}
		}

		/// <summary>
		/// Sun Apparent Longitude (deg)
		/// (Spreadsheet Column P)
		/// </summary>
		public double SunApparentLongitude
		{
			get
			{
				double returnValue = 0.0;

				returnValue = this.SunTrueLongitude - 0.00569 - 0.00478 * Math.Sin(Angle.ToRadians(125.04 - 1934.136 * this.JulianCentury));

				return returnValue;
			}
		}

		/// <summary>
		/// Mean Ecliptic Obliquity (degrees): Inclination of ecliptic plane w.r.t. celestial equator
		/// (Spreadsheet Column Q)
		/// </summary>
		public double MeanEclipticObliquity
		{
			get
			{
				double returnValue = 0.0;

				returnValue = 23 + (26 + ((21.448 - this.JulianCentury * (46.815 + this.JulianCentury * (0.00059 - this.JulianCentury * 0.001813)))) / 60) / 60;

				return returnValue;
			}
		}

		/// <summary>
		/// Obliquity Correction (degrees)
		/// (Spreadsheet Column R)
		/// </summary>
		public double ObliquityCorrection
		{
			get
			{
				double returnValue = 0.0;

				returnValue = this.MeanEclipticObliquity + 0.00256 * Math.Cos(Angle.ToRadians(125.04 - 1934.136 * this.JulianCentury));

				return returnValue;
			}
		}

		/// <summary>
		/// Solar Declination (Degrees): The measure of how many degrees North (positive) or South (negative) 
		/// of the equator that the sun is when viewed from the centre of the earth.  This varies from 
		/// approximately +23.5 (North) in June to -23.5 (South) in December.
		/// (Spreadsheet Column T)
		/// </summary>
		public double SolarDeclination
		{
			get
			{
				double returnValue = 0.0;

				returnValue = Angle.ToDegrees(Math.Asin(Math.Sin(Angle.ToRadians(this.ObliquityCorrection)) * Math.Sin(Angle.ToRadians(this.SunApparentLongitude))));

				return returnValue;
			}
		}

		/// <summary>
		/// Var Y
		/// (Spreadsheet Column U)
		/// </summary>
		public double VarY
		{
			get
			{
				double returnValue = 0.0;

				returnValue = Math.Tan(Angle.ToRadians(this.ObliquityCorrection / 2.0)) * Math.Tan(Angle.ToRadians(this.ObliquityCorrection / 2.0));

				return returnValue;
			}
		}

		/// <summary>
		/// Equation of Time (minutes)
		/// Accounts for changes in the time of solar noon for a given location over the course of a year. Earth's 
		/// elliptical orbit and Kepler's law of equal areas in equal times are the culprits behind this phenomenon.
		/// (Spreadsheet Column V)
		/// </summary>
		public double EquationOfTime
		{
			get
			{
				double returnValue = 0.0;

				returnValue = 4 * Angle.ToDegrees(this.VarY * Math.Sin(2 * Angle.ToRadians(this.SunGeometricMeanLongitude)) - 2 * this.EccentricityOfEarthOrbit * Math.Sin(Angle.ToRadians(this.SunMeanAnomaly)) + 4 * this.EccentricityOfEarthOrbit * this.VarY * Math.Sin(Angle.ToRadians(this.SunMeanAnomaly)) * Math.Cos(2 * Angle.ToRadians(this.SunGeometricMeanLongitude)) - 0.5 * this.VarY * this.VarY * Math.Sin(4 * Angle.ToRadians(this.SunGeometricMeanLongitude)) - 1.25 * this.EccentricityOfEarthOrbit * this.EccentricityOfEarthOrbit * Math.Sin(2 * Angle.ToRadians(this.SunMeanAnomaly)));

				return returnValue;
			}
		}

		/// <summary>
		/// HA Sunrise (degrees)
		/// (Spreadsheet Column W)
		/// </summary>
		public double HourAngleSunrise
		{
			get
			{
				double returnValue = 0.0;

				returnValue = Angle.ToDegrees(Math.Acos(Math.Cos(Angle.ToRadians((90 + this.AtmosphericRefraction))) / (Math.Cos(Angle.ToRadians(this.Latitude)) * Math.Cos(Angle.ToRadians(this.SolarDeclination))) - Math.Tan(Angle.ToRadians(this.Latitude)) * Math.Tan(Angle.ToRadians(this.SolarDeclination))));

				return returnValue;
			}
		}

		/// <summary>
		/// Solar Noon LST
		/// Defined for a given day for a specific longitude, it is the time when the sun crosses the meridian of 
		/// the observer's location. At solar noon, a shadow cast by a vertical pole will point either directly 
		/// north or directly south, depending on the observer's latitude and the time of year. 
		/// (Spreadsheet Column X)
		/// </summary>
		public DateTime SolarNoon
		{
			get
			{
				DateTime returnValue = DateTime.Now.Date;

				double dayFraction = (720 - (4 * this.Longitude) - this.EquationOfTime + (this.TimeZoneOffset * 60)) / 1440;
				returnValue = DateTime.Now.Date.Add(TimeSpan.FromDays(dayFraction));

				return returnValue;
			}
		}

		/// <summary>
		/// Sunlight Duration: The amount of time the sun is visible during the specified day.
		/// (Spreadsheet Column AA)
		/// </summary>
		public TimeSpan SunlightDuration
		{
			get
			{
				TimeSpan returnValue = TimeSpan.Zero;

				returnValue = TimeSpan.FromMinutes(8 * this.HourAngleSunrise);

				return returnValue;
			}
		}

		/// <summary>
		/// As light from the sun (or another celestial body) travels from the vacuum of space into Earth's atmosphere, the 
		/// path of the light is bent due to refraction. This causes stars and planets near the horizon to appear higher in 
		/// the sky than they actually are, and explains how the sun can still be visible after it has physically passed 
		/// beyond the horizon at sunset. See also apparent sunrise. Click here for a graph of atmospheric refraction vs. 
		/// elevation.
		/// </summary>
		public double AtmosphericRefraction
		{
			get
			{
				return _atmosphericRefraction;
			}
			set
			{
				_atmosphericRefraction = value;
			}
		}
		#endregion

		#region Obsolete Members
		[Obsolete("Use Angle.ToRadians() instead.", false)]
		public double ToRadians(double degrees)
		{
			return Angle.ToRadians(degrees);
		}

		[Obsolete("Use Angle.ToDegrees() instead.", false)]
		public double ToDegrees(double radians)
		{
			return Angle.ToDegrees(radians);
		}
		#endregion

		#region Additional Members
		/// <summary>
		/// Gets the True Solar Time in minutes. The Solar Time is defined as the time elapsed since the most recent 
		/// meridian passage of the sun. This system is based on the rotation of the Earth with respect to the sun. 
		/// A mean solar day is defined as the time between one solar noon and the next, averaged over the year.
		/// (Spreadsheet Column AB)
		/// </summary>
		public double TrueSolarTime
		{
			get
			{
				double returnValue = 0.0;

				returnValue = ExcelFormulae.Mod(this.TimePastLocalMidnight * 1440 + this.EquationOfTime + 4 * this.Longitude - 60 * this.TimeZoneOffset, 1440);

				return returnValue;
			}
		}
		#endregion
	}
}
