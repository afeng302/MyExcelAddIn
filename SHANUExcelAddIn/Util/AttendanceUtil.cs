using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace SHANUExcelAddIn.Util
{
    static class AttendanceUtil
    {
        public static List<AttendanceInfo> GetAttendanceInfoList(Microsoft.Office.Interop.Excel.Worksheet srcSheet)
        {
            int startRowIndex = 1;
            int startColIndex = 1;

            // locate name collum
            for (int i = 1; i < 10; i++)
            {
                string name = srcSheet.Cells[1, i].Value != null ? srcSheet.Cells[1, i].Value.ToString() : string.Empty;
                if (name.Trim() == "姓名")
                {
                    startColIndex = i;
                    break;
                }
            }

            // 姓名 日期  签到时间 签退时间
            List<AttendanceInfo> infoList = new List<AttendanceInfo>();

            for (startRowIndex = 2; startRowIndex < 10000; startRowIndex++)
            {
                // name
                string name = srcSheet.Cells[startRowIndex, 1].Value;
                if (string.IsNullOrWhiteSpace(name))
                {
                    break;
                }

                AttendanceInfo todayInfo = new AttendanceInfo(
                    srcSheet.Cells[startRowIndex, startColIndex].Value,
                    srcSheet.Cells[startRowIndex, startColIndex + 1].Value, 
                    srcSheet.Cells[startRowIndex, startColIndex + 2].Value,
                    srcSheet.Cells[startRowIndex, startColIndex + 3].Value);

                infoList.Add(todayInfo);
            }

            // sort
            infoList.Sort((info1, info2) =>
            {
                if (!info1.Name.Equals(info2.Name))
                {
                    return info1.Name.CompareTo(info2.Name);
                }

                return info1.Date.CompareTo(info2.Date);
            });

            return infoList;
        }


        public static List<AttendanceInfo> GetUnusalAttendance(List<AttendanceInfo> attendanceList)
        {
            AttendanceInfo yesterdayInfo = null;

            foreach (var todayInfo in attendanceList)
            {
                SetAttendanceState(todayInfo, yesterdayInfo);

                if (todayInfo.State != AttendanceState.None)
                {
                    Trace.WriteLine(string.Format("*** [{0}]  [{1}]  [{2}]", todayInfo.Name, todayInfo.State, todayInfo.ArriveTime));
                }

                yesterdayInfo = todayInfo;
            }

            List<AttendanceInfo> unsualInfoList = new List<AttendanceInfo>();
            attendanceList.ForEach(x
                =>
            {
                if (x.State != AttendanceState.None)
                {
                    unsualInfoList.Add(x);
                }
            });

            SetLeftState4AllPerson(unsualInfoList);

            return unsualInfoList;
        }

        public static List<PersonInfo> GetNoShowPersonList(List<PersonInfo> personInfoList, List<AttendanceInfo> attendanceInfoList)
        {
            HashSet<string> attendanceMap = new HashSet<string>();

            foreach (var nextInfo in attendanceInfoList)
            {
                attendanceMap.Add(nextInfo.Name);
            }

            List<PersonInfo> noShowList = new List<PersonInfo>();
            foreach (var nextInfo in personInfoList)
            {
                if (!attendanceMap.Contains(nextInfo.Name))
                {
                    noShowList.Add(nextInfo);
                }
            }

            return noShowList;
        }

        static void SetAttendanceState(AttendanceInfo todayInfo, AttendanceInfo yesterdayInfo)
        {
            // skip the weekend
            if ((todayInfo.ArriveTime.DayOfWeek == DayOfWeek.Saturday)
                || (todayInfo.ArriveTime.DayOfWeek == DayOfWeek.Sunday))
            {
                Trace.WriteLine("skip weekend");
                return;
            }


            // assure the same person
            if ((yesterdayInfo != null) && (yesterdayInfo.Name != todayInfo.Name))
            {
                Trace.WriteLine("name changed");
                yesterdayInfo = null;
            }

            // assure the date is contiguous
            if ((yesterdayInfo != null) && (todayInfo.ArriveTime.DayOfYear - yesterdayInfo.ArriveTime.DayOfYear != 1))
            {
                Trace.WriteLine("date is not contiguous");
                yesterdayInfo = null;
            }

            if (!todayInfo.IsValid)
            {
                todayInfo.State = AttendanceState.Absent;
                return;
            }

            if (todayInfo.WorkTime.Hours < 4)
            {
                todayInfo.State = AttendanceState.Absent;
                Trace.WriteLine("worktime less than 4 hours");
                return;
            }

            if ((todayInfo.WorkTime.Hours < 9) || (todayInfo.ArriveTime.Hour > 10))
            {
                if (yesterdayInfo == null)
                {
                    todayInfo.State = AttendanceState.Late;
                    return;
                }

                if (yesterdayInfo.WorkTime.Hours + todayInfo.WorkTime.Hours < 20)
                {
                    todayInfo.State = AttendanceState.Late;
                    return;
                }

                Trace.WriteLine("yesterday leave too late.");
            }
        }

        static void SetLeftState4AllPerson(List<AttendanceInfo> unsualInfo)
        {
            // the list has been sorted by date
            Dictionary<string, List<AttendanceInfo>> infoMap = new Dictionary<string, List<AttendanceInfo>>();
            foreach (var nextInfo in unsualInfo)
            {
                if (!infoMap.ContainsKey(nextInfo.Name))
                {
                    infoMap.Add(nextInfo.Name, new List<AttendanceInfo>());
                }
                infoMap[nextInfo.Name].Add(nextInfo);
            }

            // if one person kept absent more than 10 days, he may have left the company
            foreach (var nextInfoList in infoMap.Values)
            {
                SetLeftState4OnePerson(nextInfoList);
            }

        }

        static void SetLeftState4OnePerson(List<AttendanceInfo> unsualInfo)
        {
            if (unsualInfo.Count < 10)
            {
                return;
            }

            LeftStateChecker checker = new LeftStateChecker(unsualInfo.First<AttendanceInfo>());

            foreach (var nextInfo in unsualInfo)
            {

                checker.GetNextInfo(nextInfo);

                if (checker.IsStateWillChange)
                {
                    if (checker.NextState == AttendanceState.Left)
                    {
                        checker.ChangeState();
                        continue;
                    }

                    if (checker.NextState == AttendanceState.Absent)
                    {
                        SetLeftStateRange(unsualInfo, checker.FirstNode, checker.PreNode);
                        checker.ChangeState();
                    }
                }
            } // foreach (var nextInfo in unsualInfo)

            if (checker.PreState == AttendanceState.Left)
            {
                SetLeftStateRange(unsualInfo, checker.FirstNode, checker.PreNode);
            }
        }

        static void SetLeftStateRange(List<AttendanceInfo> unsualInfo, AttendanceInfo firstNode, AttendanceInfo currNode)
        {
            foreach (var nextInfo in unsualInfo)
            {
                if (nextInfo.Date.DayOfYear < firstNode.Date.DayOfYear)
                {
                    continue;
                }
                nextInfo.State = AttendanceState.Left;

                if (nextInfo == currNode)
                {
                    return; 
                }
            }
        }
    }
}
