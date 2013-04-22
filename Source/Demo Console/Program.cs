using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Innovative.SolarCalculator;

namespace Demo_Console
{
	class Program
	{
		static void Main(string[] args)
		{
			SolarTimes solarTimes = new SolarTimes();
			solarTimes.Latitude = 41.6;
			solarTimes.Longitude = -88;
		}
	}
}
