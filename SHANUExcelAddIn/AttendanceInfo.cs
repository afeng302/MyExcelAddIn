using SHANUExcelAddIn.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SHANUExcelAddIn
{
    enum AttendanceState
    {
        None = 0,   // initial state
        Normal,     // the normal state, should be paid as 1 day
        PaidLeave,  // weekend or holiday
        Leave,      // ask for leave (no pay)
        Late,       // be late or leave early
        Absent,     // be absent
        Dimission,  // leave the company
        NotOnboard, // not onboard yet
        AdditionalRecord,           // manager confirmed attendance
        AdditionalRecord_Ticket,    // manager confirmed attendance, but should be fined
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

        public AttendanceInfo(string name, string date, string arrivTime, string leaveTime, string exceptionComments) : this()
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                this.Name = name;
            }

            if (!string.IsNullOrWhiteSpace(date))
            {
                this.Date = Convert.ToDateTime(date);
            }

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(arrivTime))
            {
                this.ArriveTime = Convert.ToDateTime(string.Format("{0} {1}", date, arrivTime));
            }

            if (!string.IsNullOrWhiteSpace(date) && !string.IsNullOrWhiteSpace(leaveTime))
            {
                this.LeaveTime = Convert.ToDateTime(string.Format("{0} {1}", date, leaveTime));
            }

            this.ExceptionComments = exceptionComments;

            // leave later than 24:00
            if ((this.LeaveTime != DateTime.MinValue) && (this.LeaveTime < this.ArriveTime))
            {
                Trace.WriteLine("leave later than 24:00  : " + leaveTime);
                this.LeaveTime = this.LeaveTime.AddDays(1);
            }
        }

        public string Name { get; set; }

        /// <summary>
        /// will be used while there is no arrive time nor leave time (absence)
        /// 日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 签到时间
        /// </summary>
        public DateTime ArriveTime { get; set; }

        /// <summary>
        /// 签退时间
        /// </summary>
        public DateTime LeaveTime { get; set; }

        /// <summary>
        /// to indicate if this record is valid
        /// </summary>
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

        /// <summary>
        ///  work time in this day
        /// </summary>
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

        /// <summary>
        /// Over time
        /// </summary>
        public int OTHours
        {
            get
            {
                if (!this.IsValid)
                {
                    return 0;
                }

                // weekend or holiday
                if (!WorkdayUtil.IsWorkday(this.Date))
                {
                    return Convert.ToInt32(this.WorkTime.TotalHours);
                }

                // over 10 hours
                int hours = Convert.ToInt32((this.WorkTime - new TimeSpan(10, 0, 0)).TotalHours);

                return hours > 0 ? hours : 0;
            }
        }

        /// <summary>
        /// 结算天数 （1, 0.5）
        /// </summary>
        public double WorkDay
        {
           get
            {
                switch (this.State)
                {
                    case AttendanceState.None:
                        {
                            MessageBox.Show("The state None, it is Not ready to calculate workday");
                            return 0;
                        }
                    case AttendanceState.PaidLeave:
                        return 0;   // 
                    case AttendanceState.Normal:
                    case AttendanceState.Late:
                        return 1;   // the late will be fined sepately (if more than 3 times per month)
                    case AttendanceState.AdditionalRecord:
                    case AttendanceState.AdditionalRecord_Ticket:
                        return 1;   // always settle as a normal day
                    case AttendanceState.Leave:                    
                        {
                            if (this.WorkTime.TotalHours > 9.0)
                            {
                                return 1;
                            }
                            if (this.WorkTime.TotalHours > 4.0)
                            {
                                return 0.5;
                            }
                            return 0;
                        }
                    case AttendanceState.Absent:
                    case AttendanceState.Dimission:
                    case AttendanceState.NotOnboard:
                    case AttendanceState.NoShow:
                        return 0;
                    default:
                        MessageBox.Show("unhandled sate: " + this.State.ToString());
                        return 0;
                }
            }
        }

        /// <summary>
        /// 例外情况
        /// </summary>
        public string ExceptionComments { get; set; }

        public AttendanceState State { get; set; }
    }
}
