using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn.Util
{
    class WorkloadInfo
    {
        public string Name { get; set; }

        public int Month { get; set; }

        public int DueShowDays { get; set; }

        public int ActualShowDays { get; set; }

        public double PayStaffMonth { get; set; }

        public int LateDays { get; set; }

        public int OTHours { get; set; }
    }

    static class WorkloadUtil
    {
        public static List<WorkloadInfo> GetWorklaodList(List<AttendanceInfo> attendanceInfoList,
            List<AttendanceInfo> unsualInfoList)
        {
            Dictionary<string, List<AttendanceInfo>> attendanceNameMap = new Dictionary<string, List<AttendanceInfo>>();
            Dictionary<string, List<AttendanceInfo>> unsualNameMap = new Dictionary<string, List<AttendanceInfo>>();

            // split by name
            foreach (var nextInfo in attendanceInfoList)
            {
                if (!attendanceNameMap.ContainsKey(nextInfo.Name))
                {
                    attendanceNameMap.Add(nextInfo.Name, new List<AttendanceInfo>());
                }
                attendanceNameMap[nextInfo.Name].Add(nextInfo);
            }
            foreach (var nextInfo in unsualInfoList)
            {
                if (!unsualNameMap.ContainsKey(nextInfo.Name))
                {
                    unsualNameMap.Add(nextInfo.Name, new List<AttendanceInfo>());
                }
                unsualNameMap[nextInfo.Name].Add(nextInfo);
            }

            List<WorkloadInfo> workloadInfoList = new List<WorkloadInfo>();
            foreach (var nextName in attendanceNameMap.Keys)
            {
                // split by month
                Dictionary<DateTime, List<AttendanceInfo>> monthlyMap = SplitByMonth(attendanceNameMap[nextName]);

                foreach (var nextMonth in monthlyMap.Keys)
                {
                    WorkloadInfo workloadInfo = new WorkloadInfo();
                    workloadInfo.Name = nextName;
                    workloadInfo.Month = nextMonth.Month;
                    workloadInfo.DueShowDays = CalcDueShowDays(nextMonth);

                    SetWorkloadInfo(workloadInfo, monthlyMap[nextMonth]);

                    workloadInfoList.Add(workloadInfo);
                }
            } // foreach (var nextName in nameMap.Keys)

            return workloadInfoList;
        }

        static Dictionary<DateTime, List<AttendanceInfo>> SplitByMonth(List<AttendanceInfo> attendanceInfoList)
        {
            Dictionary<DateTime, List<AttendanceInfo>> monthlyMap = new Dictionary<DateTime, List<AttendanceInfo>>();

            foreach (var nextInfo in attendanceInfoList)
            {
                DateTime firstDay = new DateTime(nextInfo.Date.Year, nextInfo.Date.Month, 1);
                if (!monthlyMap.ContainsKey(firstDay))
                {
                    monthlyMap.Add(firstDay, new List<AttendanceInfo>());
                }
                monthlyMap[firstDay].Add(nextInfo);
            }

            return monthlyMap;
        }

        static void SetWorkloadInfo(WorkloadInfo workloadInfo, List<AttendanceInfo> attendanceInfoList)
        {
            int actualDays = 0;
            int lateDays = 0;
            int OTHours = 0;

            foreach (var nextInfo in attendanceInfoList)
            {
                // OT hours
                OTHours += nextInfo.OTHours;

                // actual show days
                if (nextInfo.State == AttendanceState.None)
                {
                    actualDays++;
                    continue;
                }

                // late days
                if (nextInfo.State == AttendanceState.Late)
                {
                    lateDays++;
                }
            }

            workloadInfo.ActualShowDays = actualDays;
            workloadInfo.LateDays = lateDays;
            workloadInfo.OTHours = OTHours;

            // pay staff month
            workloadInfo.PayStaffMonth = Math.Round((double)workloadInfo.ActualShowDays / (double)workloadInfo.DueShowDays, 2);
        }

        static Dictionary<DateTime, int> DUE_SHOW_DAYS_MAP = new Dictionary<DateTime, int>();
        static int CalcDueShowDays(DateTime firstDayOfMonth)
        {
            int days = 0;

            if (DUE_SHOW_DAYS_MAP.ContainsKey(firstDayOfMonth))
            {
                return DUE_SHOW_DAYS_MAP[firstDayOfMonth];
            }

            int month = firstDayOfMonth.Month;
            DateTime nextDay = firstDayOfMonth;

            do
            {
                if ((nextDay.DayOfWeek == DayOfWeek.Saturday)
                    || (nextDay.DayOfWeek == DayOfWeek.Sunday))
                {
                    nextDay = nextDay.AddDays(1);
                    continue;
                }

                if (HolidayUtil.IsHoliday(nextDay))
                {
                    nextDay = nextDay.AddDays(1);
                    continue;
                }

                days++;
                nextDay = nextDay.AddDays(1);
            } while (nextDay.Month == month);

            DUE_SHOW_DAYS_MAP[firstDayOfMonth] = days;

            return days;
        }

        static void CalcActualDay(List<AttendanceInfo> unsualList, out int absentDays, out int lateDays)
        {
            absentDays = 0;
            lateDays = 0;

            foreach (var nextInfo in unsualList)
            {
                if ((nextInfo.State == AttendanceState.Absent)
                    || (nextInfo.State == AttendanceState.Left)
                    || (nextInfo.State == AttendanceState.NoShow))
                {
                    absentDays++;
                }

                if (nextInfo.State == AttendanceState.Late)
                {
                    lateDays++;
                }
            }
        }
    }
}
