using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn.Util
{
    static class WorkdayUtil
    {
        static HashSet<string> HOLIDAY_2017 = new HashSet<string>();

        static HashSet<string> WORKING_WEEKEND_2017 = new HashSet<string>();

        /// <summary>
        /// From 2017 Q3
        /// </summary>
        static WorkdayUtil()
        {
            // Holidays
            HOLIDAY_2017.Add((new DateTime(2017, 10, 1)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 2)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 3)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 4)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 5)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 6)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 7)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 8)).ToShortDateString());

            // Working weekend
            WORKING_WEEKEND_2017.Add((new DateTime(2017, 9, 30)).ToShortDateString());
        }

        static bool IsHoliday(DateTime date)
        {
            DateTime myDate = new DateTime(date.Year, date.Month, date.Day);

            return HOLIDAY_2017.Contains(myDate.ToShortDateString());
        }

        static bool IsWorkingWeekend(DateTime date)
        {
            DateTime myDate = new DateTime(date.Year, date.Month, date.Day);

            return WORKING_WEEKEND_2017.Contains(myDate.ToShortDateString());
        }

        public static bool IsWorkday(DateTime date)
        {
            if (IsHoliday(date))
            {
                return false;
            }

            if (IsWorkingWeekend(date))
            {
                return true;
            }

            if((date.DayOfWeek == DayOfWeek.Saturday)
                || (date.DayOfWeek == DayOfWeek.Sunday))
            {
                return false;
            }

            return true;
        }
    }
}
