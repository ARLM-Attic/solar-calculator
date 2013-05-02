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
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Innovative.SolarCalculator.Tests
{
	[TestClass]
	public class ExcelFormulaeTests
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
	}
}