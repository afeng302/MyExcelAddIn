using SHANUExcelAddIn.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn
{
    enum AttendanceState
    {
        None = 0,
        PayLeave,   // weekend or holiday
        Late,       // be late or leave early
        Absent,     // be absent
        Left,       // has left
        NoShow      // no show in the attendance history
    }

    class AttendanceInfo
    {
        public AttendanceInfo()
        {
            this.ArriveTime = DateTime.MinValue;
            this.LeaveTime = DateTime.MinValue;
            this.State = AttendanceState.None;
        }

        public AttendanceInfo(string name, string date, string arriTime, string leaveTime) : this()
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                this.Name = name;
            }

            if (!string.IsNullOrWhiteSpace(date))
            {
                this.Date = Convert.ToDateTime(date);
            }

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(arriTime))
            {
                this.ArriveTime = Convert.ToDateTime(string.Format("{0} {1}", date, arriTime));
            }

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(leaveTime))
            {
                this.LeaveTime = Convert.ToDateTime(string.Format("{0} {1}", date, leaveTime));
            }

            // leave later than 24:00
            if ((this.LeaveTime != DateTime.MinValue) && (this.LeaveTime < this.ArriveTime))
            {
                Trace.WriteLine("leave later than 24:00   [{0}]", leaveTime);
                this.LeaveTime = this.LeaveTime.AddDays(1);
            }
        }

        public string Name { get; set; }

        /// <summary>
        /// will be used while there is no arrive time nor leave time (absence)
        /// </summary>
        public DateTime Date { get; set; }

        public DateTime ArriveTime { get; set; }

        public DateTime LeaveTime { get; set; }

        public bool IsValid
        {
            get
            {
                if ((this.ArriveTime == DateTime.MinValue) || (this.LeaveTime == DateTime.MinValue))
                {
                    return false;
                }

                return true;
            }
        }

        public TimeSpan WorkTime
        {
            get
            {
                if (!this.IsValid)
                {
                    return TimeSpan.MinValue;
                }

                // 王杰	2017/9/14	9:03:59	18:03:19
                // add one minute for leave to avoid the second alignment problem
                return this.LeaveTime.AddMinutes(1) - this.ArriveTime;
            }
        }

        public int OTHours
        {
            get
            {
                if (!this.IsValid)
                {
                    return 0;
                }

                // weekend or holiday
                if ((this.Date.DayOfWeek == DayOfWeek.Saturday)
                    || (this.Date.DayOfWeek == DayOfWeek.Sunday)
                    || HolidayUtil.IsHoliday(this.Date))
                {
                    return Convert.ToInt32(this.WorkTime.TotalHours);
                }

                // over 10 hours
                int hours = Convert.ToInt32((this.WorkTime - new TimeSpan(10, 0, 0)).TotalHours);

                return hours > 0 ? hours : 0;
            }
        }

        public AttendanceState State { get; set; }
    }
}
