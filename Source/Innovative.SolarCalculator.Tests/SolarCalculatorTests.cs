// ***
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
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Innovative.SolarCalculator.Tests
{
	[TestClass]
	public class SolarCalculatorTests
	{
		private TestContext _testContext = null;

		public TestContext TestContext
		{
			get
			{
				return _testContext;
			}
			set
			{
				_testContext = value;
			}
		}

		#region C#/Excel Comparison Tests
		[TestMethod]
		[TestCategory("C#/Excel Comparison Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void CsharpExcelSineComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["SIN"]);

			double actualValue = Math.Sin(value1);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.ExcelDoubleDelta, string.Format("The SIN calculation in C# does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("C#/Excel Comparison Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void CsharpExcelASineComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double sin = Convert.ToDouble(this.TestContext.DataRow["SIN"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["ASIN"]);

			double actualValue = Math.Asin(sin);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.ExcelDoubleDelta, string.Format("The ASIN calculation in C# does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("C#/Excel Comparison Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void CsharpExcelCosineComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double value2 = Convert.ToDouble(this.TestContext.DataRow["VALUE2"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["COS"]);

			double actualValue = Math.Cos(value1);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.ExcelDoubleDelta, string.Format("The COS calculation in C# does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("C#/Excel Comparison Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void CsharpExcelACosineComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double value2 = Convert.ToDouble(this.TestContext.DataRow["VALUE2"]);
			double cos = Convert.ToDouble(this.TestContext.DataRow["COS"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["ACOS"]);

			double actualValue = Math.Acos(cos);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.ExcelDoubleDelta, string.Format("The ACOS calculation in C# does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("C#/Excel Comparison Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void CsharpExcelTangentComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double value2 = Convert.ToDouble(this.TestContext.DataRow["VALUE2"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["TAN"]);

			double actualValue = Math.Tan(value1);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.ExcelDoubleDelta, string.Format("The TAN calculation in C# does not match Excel. The difference is {0}", difference));
		}
		#endregion

		#region Solar Calculation Tests
		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void JulianDayComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["JulianDay"]);
			double actualValue = solarTimes.JulianDay;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Julian Date (Column F) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void JulianCenturyComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["JulianCentury"]);
			double actualValue = solarTimes.JulianCentury;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The JulianCentury (Column G) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunGeometricMeanLongitudeComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["GeomMeanLongSun"]);
			double actualValue = solarTimes.SunGeometricMeanLongitude;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Julian Century (Column I) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunMeanAnomalyComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["GeomMeanAnomSun"]);
			double actualValue = solarTimes.SunMeanAnomaly;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Sun Mean Anomaly (Column J) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void EccentricityOfEarthOrbitComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["EccentEarthOrbit"]);
			double actualValue = solarTimes.EccentricityOfEarthOrbit;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Eccentricity Of Earth Orbit (Column K) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunEquationOfCenterComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["SunEqofCtr"]);
			double actualValue = solarTimes.SunEquationOfCenter;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Equation Of Time (Column L) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunTrueLongitudeComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["SunTrueLong"]);
			double actualValue = solarTimes.SunTrueLongitude;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Sun True Longitude (Column M) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunApparentLongitudeComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["SunAppLong"]);
			double actualValue = solarTimes.SunApparentLongitude;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Sun Apparent Longitude (Column P) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void MeanEclipticObliquityComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["MeanObliqEcliptic"]);
			double actualValue = solarTimes.MeanEclipticObliquity;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Mean Ecliptic Obliquity (Column Q) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void ObliquityCorrectionComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["ObliqCorr"]);
			double actualValue = solarTimes.ObliquityCorrection;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Obliquity Correction (Column R) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SolarDeclinationComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["SunDeclin"]);
			double actualValue = solarTimes.SolarDeclination;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Solar Declination (Column T) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void VarYComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["vary"]);
			double actualValue = solarTimes.VarY;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Var Y (Column U) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void EquationOfTimeComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["EqofTime"]);
			double actualValue = solarTimes.EquationOfTime;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Equation Of Time (Column V) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void HourAngleSunriseComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["HaSunrise"]);
			double actualValue = solarTimes.HourAngleSunrise;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The Hour Angle Sunrise (Column W) calculation does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SolarNoonComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			DateTime expectedValue = Convert.ToDateTime(this.TestContext.DataRow["SolarNoon"]);
			DateTime actualValue = solarTimes.SolarNoon;
			TimeSpan difference = expectedValue.TimeOfDay.Subtract(actualValue.TimeOfDay);

			if (difference > TestDirector.TimeSpanDelta)
			{
				Assert.Fail(string.Format("The Solar Noon (Column X) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunriseComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			DateTime expectedValue = Convert.ToDateTime(this.TestContext.DataRow["SunriseTime"]);
			DateTime actualValue = solarTimes.Sunrise;
			TimeSpan difference = expectedValue.TimeOfDay.Subtract(actualValue.TimeOfDay);

			if (difference > TestDirector.TimeSpanDelta)
			{
				Assert.Fail(string.Format("The Sunrise (Column Y) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunsetComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			DateTime expectedValue = Convert.ToDateTime(this.TestContext.DataRow["SunsetTime"]);
			DateTime actualValue = solarTimes.Sunset;
			TimeSpan difference = expectedValue.TimeOfDay.Subtract(actualValue.TimeOfDay);

			if (difference > TestDirector.TimeSpanDelta)
			{
				Assert.Fail(string.Format("The Sunset (Column Z) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void SunlightDurationComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			TimeSpan expectedValue = TimeSpan.FromMinutes(Convert.ToDouble(this.TestContext.DataRow["SunlightDuration"]));
			TimeSpan actualValue = solarTimes.SunlightDuration;
			TimeSpan difference = expectedValue - actualValue;

			if (difference > TestDirector.TimeSpanDelta)
			{
				Assert.Fail(string.Format("The Sunlight Duration (Column AA) calculation does not match Excel. The difference is {0}", difference));
			}
		}

		[TestMethod]
		[TestCategory("Solar Calculation Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "Calculations$", DataAccessMethod.Sequential)]
		public void TrueSolarTimeComparisonTest()
		{
			SolarTimes solarTimes = TestDirector.SolarTimesInstance(this.TestContext.DataRow);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["TrueSolarTime"]);
			double actualValue = solarTimes.TrueSolarTime;
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.SolarDoubleDelta, string.Format("The True Solar Time (Column AB) calculation does not match Excel. The difference is {0}", difference));
		}
		#endregion

		#region Custom Math Method Tests
		[TestMethod]
		[TestCategory("Custom Math Method Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void ToRadiansTest()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["RADIANS"]);
			double actualValue = Angle.ToRadians(value1);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.MathDoubleDelta, string.Format("ToRadians() does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Custom Math Method Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void ToDegreesTest()
		{
			double radians = Convert.ToDouble(this.TestContext.DataRow["VALUE3"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["DEGREES"]);
			double actualValue = Angle.ToDegrees(radians);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.MathDoubleDelta, string.Format("ToDegrees() does not match Excel. The difference is {0}", difference));
		}
		#endregion

		#region Excel Formulae Tests
		[TestMethod]
		[TestCategory("Excel Formulae Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "ExcelFormulas$", DataAccessMethod.Sequential)]
		public void ExcelModuloComparisons()
		{
			double value1 = Convert.ToDouble(this.TestContext.DataRow["VALUE1"]);
			double value2 = Convert.ToDouble(this.TestContext.DataRow["VALUE2"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["MOD"]);

			double actualValue = ExcelFormulae.Mod(value1, value2);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.MathDoubleDelta, string.Format("The Modulo (% operator) calculation in C# does not match Excel. The difference is {0}", difference));
		}

		[TestMethod]
		[TestCategory("Excel Formulae Tests")]
		[DeploymentItem("NOAA Solar Calculations Test Data.xlsx")]
		[DataSource("System.Data.OleDb", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=NOAA Solar Calculations Test Data.xlsx;Extended Properties=\"Excel 12.0;HDR=YES\"", "DateValue$", DataAccessMethod.Sequential)]
		public void ExcelDateValueTest()
		{
			DateTime value1 = Convert.ToDateTime(this.TestContext.DataRow["DATE"]);
			double expectedValue = Convert.ToDouble(this.TestContext.DataRow["DATEVALUE"]);

			double actualValue = ExcelFormulae.ToExcelDateValue(value1);
			double difference = expectedValue - actualValue;

			Assert.AreEqual(expectedValue, actualValue, TestDirector.MathDoubleDelta, string.Format("ToRadians() does not match Excel. The difference is {0}", difference));
		}
		#endregion
	}
}
