using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn
{
    class AttendanceInfo
    {
        public string Name { get; set; }

        public DateTime ArriveTime { get; set; }

        public DateTime LeaveTime { get; set; }

        public TimeSpan WorkTime { get; set; }
    }
}
