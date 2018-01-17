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

        public double ActualShowDays { get; set; }

        public double PayStaffMonth { get; set; }

        public int AbsentDays { get; set; }

        public int LateTimes { get; set; }

        public int LeaveTimes { get; set; }

        public int AdditionalRecordTimes { get; set; }

        public int OTHours { get; set; }
    }

    static class WorkloadUtil
    {
        public static List<WorkloadInfo> GetWorklaodListPerMonth(List<AttendanceInfo> attendanceInfoList)
        {
            Dictionary<string, List<AttendanceInfo>> attendanceNameMap = new Dictionary<string, List<AttendanceInfo>>();

            // group by name
            foreach (var nextInfo in attendanceInfoList)
            {
                if (!attendanceNameMap.ContainsKey(nextInfo.Name))
                {
                    attendanceNameMap.Add(nextInfo.Name, new List<AttendanceInfo>());
                }
                attendanceNameMap[nextInfo.Name].Add(nextInfo);
            }

            List<WorkloadInfo> workloadInfoList = new List<WorkloadInfo>();
            foreach (var nextName in attendanceNameMap.Keys)
            {
                // group by month
                Dictionary<DateTime, List<AttendanceInfo>> monthlyMap = GroupByMonth(attendanceNameMap[nextName]);

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

        public static List<WorkloadInfo> GetWorkloadListTotally(List<WorkloadInfo> workloadPerMonth)
        {
            Dictionary<string, List<WorkloadInfo>> workloadMap = new Dictionary<string, List<WorkloadInfo>>();

            // group by name
            foreach (var nextInfo in workloadPerMonth)
            {
                if (!workloadMap.ContainsKey(nextInfo.Name))
                {
                    workloadMap[nextInfo.Name] = new List<WorkloadInfo>();
                }
                workloadMap[nextInfo.Name].Add(nextInfo);
            }

            // summarize by name
            List<WorkloadInfo> totalInfo = new List<WorkloadInfo>();
            foreach (var nextPersonInfoList in workloadMap.Values)
            {
                WorkloadInfo firstInfo = nextPersonInfoList.First();

                // to advoid the precision difference after summaried, we round two decimal before sum
                firstInfo.PayStaffMonth = Math.Round(firstInfo.PayStaffMonth, 2);

                // deduce money by late days (per month)
                // we escape the latedays property for absent days in summarizatioin
                if (firstInfo.LateTimes >= 3)
                {
                    firstInfo.LateTimes = (firstInfo.LateTimes + 2) / 3;
                }
                else
                {
                    // reset the late days
                    firstInfo.LateTimes = 0;
                }

                foreach (var nextPersonInfo in nextPersonInfoList)
                {
                    if (nextPersonInfo == firstInfo)
                    {
                        continue;
                    }
                    firstInfo.ActualShowDays += nextPersonInfo.ActualShowDays;
                    firstInfo.DueShowDays += nextPersonInfo.DueShowDays;
                    if (nextPersonInfo.LateTimes >= 3)
                    {
                        firstInfo.LateTimes += ((nextPersonInfo.LateTimes + 2) / 3);
                    }
                    firstInfo.OTHours += nextPersonInfo.OTHours;
                    firstInfo.PayStaffMonth += Math.Round(nextPersonInfo.PayStaffMonth, 2);
                }
                totalInfo.Add(firstInfo);
            }

            return totalInfo;
        }

        static Dictionary<DateTime, List<AttendanceInfo>> GroupByMonth(List<AttendanceInfo> attendanceInfoList)
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
            double actualDays = 0;
            int lateTimes = 0;
            int leaveTimes = 0;
            int additionalRecordTimes = 0;
            int OTHours = 0;

            foreach (var nextInfo in attendanceInfoList)
            {
                // OT hours
                OTHours += nextInfo.OTHours;

                // actural show days
                actualDays += nextInfo.WorkDay;

                // late times
                if (nextInfo.State == AttendanceState.Late)
                {
                    lateTimes++;
                }

                // leave times
                if (nextInfo.State == AttendanceState.Leave)
                {
                    leaveTimes++;
                }

                // additional record times
                if ((nextInfo.State == AttendanceState.AdditionalRecord)
                    || (nextInfo.State == AttendanceState.AdditionalRecord_Ticket))
                {
                    additionalRecordTimes++;
                }
            }

            workloadInfo.ActualShowDays = actualDays;
            workloadInfo.LateTimes = lateTimes;
            workloadInfo.LeaveTimes = leaveTimes;
            workloadInfo.AdditionalRecordTimes = additionalRecordTimes;
            workloadInfo.OTHours = OTHours;

            // pay staff month
            workloadInfo.PayStaffMonth = workloadInfo.ActualShowDays / (double)workloadInfo.DueShowDays;
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
                // skip the dayoff
                if (!WorkdayUtil.IsWorkday(nextDay))
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

        //static void CalcActualDay(List<AttendanceInfo> unsualList, out int absentDays, out int lateDays)
        //{
        //    absentDays = 0;
        //    lateDays = 0;

        //    foreach (var nextInfo in unsualList)
        //    {
        //        if ((nextInfo.State == AttendanceState.Absent)
        //            || (nextInfo.State == AttendanceState.Dimission)
        //            || (nextInfo.State == AttendanceState.NoShow))
        //        {
        //            absentDays++;
        //        }

        //        if (nextInfo.State == AttendanceState.Late)
        //        {
        //            lateDays++;
        //        }
        //    }
        //}
    }
}
