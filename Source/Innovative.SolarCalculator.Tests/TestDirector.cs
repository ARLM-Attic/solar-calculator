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
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;

namespace Innovative.SolarCalculator.Tests
{
    public class TestDirector
    {
		private static IEnumerable<DataRow> _mathTestData = null;
		private static IEnumerable<DataRow> _solarTestData = null;
		private static IEnumerable<DataRow> _oleAutomationTestData = null;
        private static SolarTimes _solarTimes = null;

        /// <summary>
        /// Specifies the precision when comparing data values of type double
        /// for math comparison tests.
        /// </summary>
		public static double MathDoubleDelta
		{
			get
			{
				return 0.000000000001;
			}
		}

        /// <summary>
        /// Specifies the precision when comparing data values of type double
        /// for Excel comparison tests.
        /// </summary>
        public static double ExcelDoubleDelta
        {
            get
            {
                return 0.000000000001;
            }
        }

        /// <summary>
        /// Specifies the precision when comparing data values of type double
        /// for solar comparison tests.
        /// </summary>
        public static double SolarDoubleDelta
        {
            get
            {
                return 0.00000000001;
            }
        }

        /// <summary>
        /// Creates a SolarTimes instance for use in test cases.
        /// </summary>
        public static SolarTimes SolarTimesInstance
        {
            get
            {
                if (_solarTimes == null)
                {
                    _solarTimes = new SolarTimes();
                    _solarTimes.Latitude = Properties.Settings.Default.Latitude;
                    _solarTimes.Longitude = Properties.Settings.Default.Longitude;
                }

                return _solarTimes;
            }
        }

        /// <summary>
        /// Gets the test data required for the mathmatical tests.
        /// </summary>
		public static IEnumerable<DataRow> MathTestData
		{
			get
			{
				if (_mathTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [ExcelFormulas$]";
					DataSet data = excel.ExecuteDataSet(sql, "ExcelFormulas");
					_mathTestData = data.Tables["ExcelFormulas"].AsEnumerable();
				}

				return _mathTestData;
			}
		}

        /// <summary>
        /// Gets the test data required for the solar tests.
        /// </summary>
		public static IEnumerable<DataRow> SolarTestData
		{
			get
			{
				if (_solarTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [Calculations$]";
                    DataSet data = excel.ExecuteDataSet(sql, "ExcelFormulas");
                    _solarTestData = data.Tables["ExcelFormulas"].AsEnumerable();
				}

				return _solarTestData;
			}
		}

        /// <summary>
        /// Gets the test data required for the DateTime extension tests.
        /// </summary>
		public static IEnumerable<DataRow> DateValueTestData
		{
			get
			{
				if (_oleAutomationTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [DateValue$]";
                    DataSet data = excel.ExecuteDataSet(sql, "ExcelFormulas");
                    _oleAutomationTestData = data.Tables["ExcelFormulas"].AsEnumerable();
				}

				return _oleAutomationTestData;
			}
		}
    }
}
