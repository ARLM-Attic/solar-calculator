using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if NET20
namespace Innovative.SolarCalculator.Net20.Tests
#elif NET35
namespace Innovative.SolarCalculator.Net35.Tests
#elif NET40
namespace Innovative.SolarCalculator.Net40.Tests
#elif NET45
namespace Innovative.SolarCalculator.Net45.Tests
#elif NET451
namespace Innovative.SolarCalculator.Net451.Tests
#elif PORTABLE40
namespace Innovative.SolarCalculator.Portable40.Tests
#elif PORTABLE45
namespace Innovative.SolarCalculator.Portable40.Tests
#endif
{
	public static class CustomAssert
	{
		public static void AreEqual(decimal expected, decimal actual, decimal delta)
		{
			decimal difference = Math.Abs(expected) - Math.Abs(actual);

			if (difference > delta)
			{
				string message = string.Format("Expected  {0}, Actual = {1}, Difference = {2} which is greater than the delta of {3}", expected, actual, difference, delta);
				Assert.Fail(message);
			}
		}
	}
}
