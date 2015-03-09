using System;

namespace EmailSender.Shared
{
	// TODO Think of a better name for this class
	public static class Helper
	{
		// Increases array index, when limit is reached next value is 0 (zero)
		public static int IncreaseIndex(int index, int limit)
		{
			int incIndex;
			if (index == limit - 1)
			{
				incIndex = 0;
			}
			else
			{
				incIndex = index + 1;
			}

			return incIndex;
		}

		// Gets date for scheduling meetings or creating tasks (1-3 days from today, at 9-16 o'clock)
		public static DateTime GetDateForScheduling(DateTime currentTime, Random rnd)
		{
			DateTime date = currentTime.AddDays(rnd.Next(1, 3));

			return new DateTime(date.Year, date.Month, date.Day, rnd.Next(9, 17), 0, 0);
		}
	}
}
