using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Diagnostics;

namespace Innovative.SolarCalculator.Tests
{
	[TestClass]
	public class SolarCalculatorTests
	{
		#region C#/Excel Comparison Tests
		[TestMethod]
		public void CsharpExcelModuloComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double value2 = item.Field<double>("VALUE2");
				double expectedValue = item.Field<double>("MOD");

				int actualValue = (int)value1 % (int)value2;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("Modulo calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void CsharpExcelSineComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double expectedValue = item.Field<double>("SIN");

				double actualValue = Math.Sin(value1);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("SIN calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void CsharpExcelASineComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double sin = item.Field<double>("SIN");
				double expectedValue = item.Field<double>("ASIN");

				double actualValue = Math.Asin(sin);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("ASIN calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void CsharpExcelCosineComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double value2 = item.Field<double>("VALUE2");
				double expectedValue = item.Field<double>("COS");

				double actualValue = Math.Cos(value1);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("COS calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void CsharpExcelACosineComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double value2 = item.Field<double>("VALUE2");
				double cos = item.Field<double>("COS");
				double expectedValue = item.Field<double>("ACOS");

				double actualValue = Math.Acos(cos);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("ACOS calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void CsharpExcelTangentComparisons()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double value2 = item.Field<double>("VALUE2");
				double expectedValue = item.Field<double>("TAN");

				double actualValue = Math.Tan(value1);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("TAN calculation in C# does not match Excel. The difference is {0}", difference));
			}
		}
		#endregion

		#region Solar Calculation Tests
		[TestMethod]
		public void JulianDayComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("JulianDay");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.JulianDay;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Julian Date (Column F)calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void JulianCenturyComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("JulianCentury");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.JulianCentury;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The JulianCentury (Column G)calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunGeometricMeanLongitudeComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("GeomMeanLongSun");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SunGeometricMeanLongitude;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Julian Century (Column I) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunMeanAnomalyComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("GeomMeanAnomSun");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SunMeanAnomaly;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Sun Mean Anomaly (Column J) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void EccentricityOfEarthOrbitComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("EccentEarthOrbit");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.EccentricityOfEarthOrbit;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Eccentricity Of Earth Orbit (Column K) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunEquationOfCenterComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("SunEqofCtr");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SunEquationOfCenter;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Equation Of Time (Column L) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunTrueLongitudeComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("SunTrueLong");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SunTrueLongitude;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Sun True Longitude (Column M) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunApparentLongitudeComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("SunAppLong");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SunApparentLongitude;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Sun Apparent Longitude (Column P) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void MeanEclipticObliquityComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("MeanObliqEcliptic");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.MeanEclipticObliquity;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Mean Ecliptic Obliquity (Column Q) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void ObliqCorrComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("ObliqCorr");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.ObliqCorr;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Obliq Corr (Column R) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SolarDeclinationComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("SunDeclin");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.SolarDeclination;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Solar Declination (Column T) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void VarYComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("vary");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.VarY;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Var Y (Column U) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void EquationOfTimeComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("EqofTime");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.EquationOfTime;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Equation Of Time (Column V) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void HaSunriseComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				double expectedValue = item.Field<double>("HaSunrise");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				double actualValue = TestDirector.SolarTimesInstance.HaSunrise;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The HA Sunrise (Column W) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SolarNoonComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				DateTime expectedValue = item.Field<DateTime>("SolarNoon");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				DateTime actualValue = TestDirector.SolarTimesInstance.SolarNoon;
				TimeSpan difference = expectedValue.Subtract(actualValue);

				Assert.AreEqual(expectedValue.TimeOfDay, actualValue.TimeOfDay, string.Format("The Solar Noon (Column X) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunriseComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				DateTime expectedValue = item.Field<DateTime>("SunriseTime");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				DateTime actualValue = TestDirector.SolarTimesInstance.Sunrise;
				TimeSpan difference = expectedValue.Subtract(actualValue);

				Assert.AreEqual(expectedValue.TimeOfDay, actualValue.TimeOfDay, string.Format("The Sunrise (Column Y) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunsetComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				DateTime expectedValue = item.Field<DateTime>("SunsetTime");

				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);
				DateTime actualValue = TestDirector.SolarTimesInstance.Sunset;
				TimeSpan difference = expectedValue.TimeOfDay.Subtract(actualValue.TimeOfDay);

				Assert.AreEqual(expectedValue.TimeOfDay, actualValue.TimeOfDay, string.Format("The Sunset (Column Z) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void SunlightDurationComparisonTest()
		{
			foreach (var item in TestDirector.SolarTestData)
			{
				DateTime date = item.Field<DateTime>("Date");
				DateTime time = item.Field<DateTime>("Time");
				TestDirector.SolarTimesInstance.ForDate = date.Add(time.TimeOfDay);

				double expectedValue = item.Field<double>("SunlightDuration");				
				double actualValue = TestDirector.SolarTimesInstance.SunlightDuration.TotalMinutes;
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("The Sunlight Duration (Column AA) calculation does not match Excel. The difference is {0}", difference));
			}
		}
		#endregion

		#region Custom Math Method Tests
		[TestMethod]
		public void ExcelDateValueTest()
		{
			foreach (var item in TestDirector.DateValueTestData)
			{
				DateTime value1 = item.Field<DateTime>("DATE");
				double expectedValue = item.Field<double>("DATEVALUE");

				double actualValue = value1.ToExcelDateValue();
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("ToRadians() does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void ToRadiansTest()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double value1 = item.Field<double>("VALUE1");
				double expectedValue = item.Field<double>("RADIANS");

				double actualValue = TestDirector.SolarTimesInstance.ToRadians(value1);
				double difference = expectedValue - actualValue;
				
				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("ToRadians() does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		public void ToDegreesTest()
		{
			foreach (var item in TestDirector.MathTestData)
			{
				double radians = item.Field<double>("RADIANS");
				double expectedValue = item.Field<double>("DEGREES");

				double actualValue = TestDirector.SolarTimesInstance.ToDegrees(radians);
				double difference = expectedValue - actualValue;

				Assert.AreEqual(expectedValue, actualValue, TestDirector.DoubleDelta, string.Format("ToDegrees() does not match Excel. The difference is {0}", difference));
			}
		}
		#endregion
	}
}
