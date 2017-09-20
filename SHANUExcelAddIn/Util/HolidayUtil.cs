using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn.Util
{
    static class HolidayUtil
    {
        static HashSet<string> HOLIDAY_2017 = new HashSet<string>();

        static HolidayUtil()
        {
            // National Holiday
            HOLIDAY_2017.Add((new DateTime(2017, 10, 1)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 2)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 3)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 4)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 5)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 6)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 7)).ToShortDateString());
            HOLIDAY_2017.Add((new DateTime(2017, 10, 8)).ToShortDateString());
        }

        public static bool IsHoliday(DateTime date)
        {
            return HOLIDAY_2017.Contains(date.ToShortDateString());
        }
    }
}
