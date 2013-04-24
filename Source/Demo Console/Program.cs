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
			SolarTimes solarTimes = new SolarTimes(DateTime.Now);

			string daylightName = TimeZoneInfo.Local.DaylightName;
			bool isSavings = TimeZoneInfo.Local.IsDaylightSavingTime(DateTime.Now);
			double offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).TotalHours;
			TimeZoneInfo.AdjustmentRule[] rules = TimeZoneInfo.Local.GetAdjustmentRules();

			// ***
			// *** Geo coordinates
			// ***
			solarTimes.Latitude = 41.6042880444544;
			solarTimes.Longitude = -88.03034663200378;

			// ***
			// *** Display the sunrise and sunset
			// ***
			Console.WriteLine("Sunrise is at {0} on {1}.", solarTimes.Sunrise.ToLongTimeString(), solarTimes.Sunrise.ToLongDateString());
			Console.WriteLine("Sunset is at {0} on {1}.", solarTimes.Sunset.ToLongTimeString(), solarTimes.Sunset.ToLongDateString());
			Console.WriteLine("The amount of sunlight on {0} is {1} hours, {2} minutes and {3} seconds.", solarTimes.ForDate.ToLongDateString(), solarTimes.SunlightDuration.Hours, solarTimes.SunlightDuration.Minutes, solarTimes.SunlightDuration.Seconds);
		}
	}
}
