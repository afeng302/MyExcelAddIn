using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn.Util
{
    class LeftStateChecker
    {
        public const int MAX_ABSENT_DAYS = 22; // when exceed the days will be treated as left

        public AttendanceInfo FirstNode { get; private set; }
        
        public AttendanceInfo PreNode { get; private set; }

        public AttendanceInfo NextNode { get; private set; }

        public int AbsentDays { get; private set; }

        public AttendanceState PreState { get; private set; }

        public AttendanceState NextState { get; private set; }

        public LeftStateChecker(AttendanceInfo firstNode)
        {
            this.FirstNode = firstNode;
            this.PreNode = firstNode;

            this.AbsentDays = 0;

            this.PreState = AttendanceState.Absent;

            this.NextState = AttendanceState.Absent;
        }

        public void GetNextInfo(AttendanceInfo nextInfo)
        {
            // assert there is no weekend days in the list
            if ((nextInfo.Date.DayOfWeek == DayOfWeek.Saturday)
                || (nextInfo.Date.DayOfWeek == DayOfWeek.Sunday))
            {
                Debug.Assert(false, "there is weekend in the list");
            }

            // filter out the absent days
            if (nextInfo.State != AttendanceState.Absent)
            {
                return;
            }

            // first node
            if (this.AbsentDays == 0)
            {
                this.Reset(nextInfo);

                this.AbsentDays++;

                return;
            }

            // date is not continous, reset the points
            if ((nextInfo.Date.DayOfWeek == DayOfWeek.Monday)
                && (nextInfo.Date.DayOfYear - this.PreNode.Date.DayOfYear > 3))
            {
                if (this.PreState == AttendanceState.Left)
                {
                    // change to absent
                    this.NextState = AttendanceState.Absent;
                    this.NextNode = nextInfo;
                }
                else
                {
                    this.Reset(nextInfo);
                }

                return;
            }
            if ((nextInfo.Date.DayOfWeek != DayOfWeek.Monday)
               && (nextInfo.Date.DayOfYear - this.PreNode.Date.DayOfYear > 1))
            {
                if (this.PreState == AttendanceState.Left)
                {
                    // change to absent
                    this.NextState = AttendanceState.Absent;
                    this.NextNode = nextInfo;
                }
                else
                {
                    this.Reset(nextInfo);
                }
                return;
            }

            // change state to left
            if ((++this.AbsentDays > MAX_ABSENT_DAYS) && (this.PreState == AttendanceState.Absent))
            {
                this.NextState = AttendanceState.Left;
                this.NextNode = nextInfo;
            }

            this.PreNode = nextInfo;
        }

        public void Reset(AttendanceInfo currNode)
        {
            this.AbsentDays = 0;

            this.PreNode = currNode;
            this.FirstNode = currNode;

            this.PreState = AttendanceState.Absent;

            this.NextState = AttendanceState.Absent;
        }

        public bool IsStateWillChange
        {
            get
            {
                return this.PreState != this.NextState;
            }
        }

        public void ChangeState()
        {
            // reset the days
            if (this.NextState == AttendanceState.Absent)
            {
                this.AbsentDays = 0; 
            }

            this.PreState = this.NextState;
            this.PreNode = this.NextNode;
        }
    }
}
