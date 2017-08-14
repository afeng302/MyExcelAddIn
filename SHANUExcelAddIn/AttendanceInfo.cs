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
        Late,   // be late or leave early
        Absent  // no show
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

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(arriTime))
            {
                this.ArriveTime = Convert.ToDateTime(string.Format("{0} {1}", date, arriTime));
            }

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(leaveTime))
            {
                this.LeaveTime = Convert.ToDateTime(string.Format("{0} {1}", date, leaveTime));
            }

            // leave later than 24:00
            if (this.LeaveTime < this.ArriveTime)
            {
                Trace.WriteLine("leave later than 24:00   [{0}]", leaveTime);
                this.LeaveTime = this.LeaveTime.AddDays(1);
            }
        }

        public string Name { get; set; }

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

                return this.LeaveTime - this.ArriveTime;
            }
        }

        public AttendanceState State { get; set; }
    }
}
