using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn.Util
{
    static class WorkdayUtil
    {
        static HashSet<string> HOLIDAY = new HashSet<string>();

        static HashSet<string> WORKING_WEEKEND = new HashSet<string>();

        /// <summary>
        /// From 2017 Q3
        /// </summary>
        static WorkdayUtil()
        {
            //
            // Holidays
            #region 2017
            HOLIDAY.Add((new DateTime(2017, 10, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 2)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 3)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 4)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 5)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 6)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 7)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2017, 10, 8)).ToShortDateString());
            #endregion // 2017
            
            #region 2018
            HOLIDAY.Add((new DateTime(2018, 1, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 15)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 16)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 17)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 18)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 19)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 20)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 2, 21)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 4, 5)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 4, 6)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 4, 7)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 4, 29)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 4, 30)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 5, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 6, 16)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 6, 17)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 6, 18)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 9, 22)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 9, 23)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 9, 24)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 2)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 3)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 4)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 5)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 6)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 10, 7)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2018, 12, 31)).ToShortDateString());
            #endregion // 2018

            #region 2019
            HOLIDAY.Add((new DateTime(2019, 1, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 2, 4)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 2, 5)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 2, 6)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 2, 7)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 2, 8)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 4, 5)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 5, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 5, 2)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 5, 3)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 6, 7)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 9, 13)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 10, 1)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 10, 2)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 10, 3)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 10, 4)).ToShortDateString());
            HOLIDAY.Add((new DateTime(2019, 10, 7)).ToShortDateString());
            #endregion // 2019

            //
            // Working weekend
            // 2017
            WORKING_WEEKEND.Add((new DateTime(2017, 9, 30)).ToShortDateString());
            // 2018
            WORKING_WEEKEND.Add((new DateTime(2018, 2, 10)).ToShortDateString()); // exchanged with the workday 2.11
            WORKING_WEEKEND.Add((new DateTime(2018, 2, 24)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2018, 4, 8)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2018, 4, 28)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2018, 9, 29)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2018, 9, 30)).ToShortDateString());
            // 2019
            WORKING_WEEKEND.Add((new DateTime(2019, 2, 2)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2019, 2, 3)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2019, 4, 28)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2019, 5, 5)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2019, 9, 29)).ToShortDateString());
            WORKING_WEEKEND.Add((new DateTime(2019, 10, 12)).ToShortDateString());
        }

        static bool IsHoliday(DateTime date)
        {
            DateTime myDate = new DateTime(date.Year, date.Month, date.Day);

            return HOLIDAY.Contains(myDate.ToShortDateString());
        }

        static bool IsWorkingWeekend(DateTime date)
        {
            DateTime myDate = new DateTime(date.Year, date.Month, date.Day);

            return WORKING_WEEKEND.Contains(myDate.ToShortDateString());
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
