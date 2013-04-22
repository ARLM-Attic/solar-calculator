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

		public static double DoubleDelta
		{
			get
			{
				return 0.0000000000000009;
			}
		}

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

		public static IEnumerable<DataRow> MathTestData
		{
			get
			{
				if (_mathTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [ExcelFormulas$]";
					DataSet data = excel.ExecuteDataSetQuery(sql);
					_mathTestData = data.Tables[0].AsEnumerable();
				}

				return _mathTestData;
			}
		}

		public static IEnumerable<DataRow> SolarTestData
		{
			get
			{
				if (_solarTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [Calculations$]";
					DataSet data = excel.ExecuteDataSetQuery(sql);
					_solarTestData = data.Tables[0].AsEnumerable();
				}

				return _solarTestData;
			}
		}

		public static IEnumerable<DataRow> DateValueTestData
		{
			get
			{
				if (_oleAutomationTestData == null)
				{
					ExcelQuery excel = new ExcelQuery("NOAA Solar Calculations Test Data.xlsx");
					string sql = "SELECT * FROM [DateValue$]";
					DataSet data = excel.ExecuteDataSetQuery(sql);
					_oleAutomationTestData = data.Tables[0].AsEnumerable();
				}

				return _oleAutomationTestData;
			}
		}
    }
}
