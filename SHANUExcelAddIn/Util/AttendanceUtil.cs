using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

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

            for (startRowIndex = 2; startRowIndex < 50000; startRowIndex++)
            {
                // name
                string name = srcSheet.Cells[startRowIndex, startColIndex].Value;
                if (string.IsNullOrWhiteSpace(name))
                {
                    if (startRowIndex > 10)
                    {
                        break;
                    }
                    continue; // the top lines may have some statistic infomation
                }

                AttendanceInfo todayInfo = new AttendanceInfo(
                    srcSheet.Cells[startRowIndex, startColIndex].Value,
                    srcSheet.Cells[startRowIndex, startColIndex + 1].Value,
                    srcSheet.Cells[startRowIndex, startColIndex + 2].Value,
                    srcSheet.Cells[startRowIndex, startColIndex + 3].Value);

                infoList.Add(todayInfo);
            }

            // pad absent days
            PadAbsentDays(infoList);

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

        /// <summary>
        /// get unsual attandence list and also set the state on origin items
        /// </summary>
        /// <param name="attendanceList"></param>
        /// <returns></returns>
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

            // correct the attendance mistake caused by system problem
            AttendanceCorrection(attendanceList);

            List<AttendanceInfo> unsualInfoList = new List<AttendanceInfo>();
            attendanceList.ForEach(x =>
            {
                if ((x.State != AttendanceState.None) && (x.State != AttendanceState.PayLeave))
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

        /// <summary>
        /// it is possible for the attandence info has no some days when the person did not punch card at all.
        /// </summary>
        /// <param name="infoList"></param>
        static void PadAbsentDays(List<AttendanceInfo> infoList)
        {
            Dictionary<string, List<AttendanceInfo>> nameMap = new Dictionary<string, List<AttendanceInfo>>();

            // sort by date to get the first and last day
            infoList.Sort((info1, info2) => info1.Date.CompareTo(info2.Date));
            DateTime firstStatisticDay = infoList.First<AttendanceInfo>().Date;
            DateTime lastStatisticDay = infoList.Last<AttendanceInfo>().Date;

            // group by name
            foreach (var nextInfo in infoList)
            {
                if (!nameMap.ContainsKey(nextInfo.Name))
                {
                    nameMap[nextInfo.Name] = new List<AttendanceInfo>();
                }
                nameMap[nextInfo.Name].Add(nextInfo);
            }

            // check duplicate name in the attendance record
            int totaldays = 0;
            foreach (var nextPersonInfoList in nameMap.Values)
            {
                totaldays += nextPersonInfoList.Count;
            }
            int averageDays = totaldays / nameMap.Count;
            foreach (var nextName in nameMap.Keys)
            {
                Debug.Assert(nameMap[nextName].Count <= averageDays, "there is duplicate names?  " + nextName);
            }

            // pad the absent days
            List<AttendanceInfo> padInfoList = new List<AttendanceInfo>();
            foreach (var nextPersonInfoList in nameMap.Values)
            {
                // sort by date
                nextPersonInfoList.Sort((info1, info2) => info1.Date.CompareTo(info2.Date));

                // pad absent days at header
                AttendanceInfo firstInfo = nextPersonInfoList.First<AttendanceInfo>();
                DateTime absentDay = firstStatisticDay;
                while (absentDay != firstInfo.Date)
                {
                    // pad absent day
                    padInfoList.Add(new AttendanceInfo()
                    {
                        Name = firstInfo.Name,
                        Date = absentDay,
                        State = AttendanceState.Absent
                    });

                    absentDay = absentDay.AddDays(1);
                }

                // pad absent days during work days
                DateTime yesterday = firstInfo.Date;
                foreach (var nextInfo in nextPersonInfoList)
                {
                    if (nextInfo.Date == yesterday)
                    {
                        continue; // first day
                    }

                    while (yesterday.AddDays(1) != nextInfo.Date)
                    {
                        // pad absent day
                        padInfoList.Add(new AttendanceInfo()
                        {
                            Name = nextInfo.Name,
                            Date = yesterday.AddDays(1),
                            State = AttendanceState.Absent
                        });
                        yesterday = yesterday.AddDays(1);
                    }

                    yesterday = nextInfo.Date;
                }

                // pad absent days at tail
                AttendanceInfo lastInfo = nextPersonInfoList.Last<AttendanceInfo>();
                absentDay = lastInfo.Date;
                while (absentDay != lastStatisticDay)
                {
                    // pad absent day
                    padInfoList.Add(new AttendanceInfo()
                    {
                        Name = lastInfo.Name,
                        Date = absentDay.AddDays(1),
                        State = AttendanceState.Absent
                    });
                    absentDay = absentDay.AddDays(1);
                }

            } // foreach (var nextPersonInfoList in nameMap.Values)

            infoList.AddRange(padInfoList);
        }

        static void SetAttendanceState(AttendanceInfo todayInfo, AttendanceInfo yesterdayInfo)
        {
            // skip the dayoff (weeken dand holiday)
            if (!WorkdayUtil.IsWorkday(todayInfo.Date))
            {
                todayInfo.State = AttendanceState.PayLeave;
                Trace.WriteLine("skip weekend: " + todayInfo.Date.ToShortDateString());
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

                if (yesterdayInfo.WorkTime.Hours + todayInfo.WorkTime.Hours < 18)
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
            if (unsualInfo.Count < LeftStateChecker.MAX_ABSENT_DAYS)
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

        /// <summary>
        /// correct the attendance mistake caused by system problem
        /// </summary>
        static void AttendanceCorrection(List<AttendanceInfo> attendanceList)
        {
            // 8.31 has no leave time record in system
            #region 8.31 has no leave time record in system
            foreach (var nextInfo in attendanceList)
            {
                if (nextInfo.Date.Equals(new DateTime(2017, 8, 31)))
                {
                    PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                    if (personInfo == null)
                    {
                        //MessageBox.Show(nextInfo.Name + " 没有台账信息!");
                        Trace.WriteLine("cannot find " + nextInfo.Name);
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(personInfo.DimissionDate))
                    {
                        // correct the status
                        nextInfo.State = AttendanceState.None;
                        continue;
                    }

                    DateTime dimissionDate = DateTime.MinValue;

                    try
                    {
                        dimissionDate = Convert.ToDateTime(personInfo.DimissionDate);
                    }
                    catch (FormatException)
                    {
                        dimissionDate = DateTime.MinValue;
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(exp.Message);
                        return;
                    }

                    if (dimissionDate < new DateTime(2017, 8, 31))
                    {
                        continue; // skip the leave early case
                    }

                    // correct the status
                    nextInfo.State = AttendanceState.None;

                } // if (nextInfo.Date.Equals(new DateTime(2017, 8, 31)))
            } // foreach (var nextInfo in attendanceList)
            #endregion // 8.31 has no leave time record in system


            //
            // befor or later than enter and leave date
            // always put this rule in the last
            // 
            #region // befor or later than enter and leave date
            foreach (var nextInfo in attendanceList)
            {
                DateTime onboardDate = DateTime.MinValue;
                DateTime dimissionDate = DateTime.MinValue;

                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                if (personInfo == null)
                {
                    //MessageBox.Show(nextInfo.Name + " 没有台账信息!");
                    Trace.WriteLine("cannot find " + nextInfo.Name);
                    continue;
                }

                // enter date
                try
                {
                    onboardDate = Convert.ToDateTime(personInfo.OnboardDate);
                }
                catch (FormatException)
                {
                    onboardDate = DateTime.MinValue;
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message);
                    return;
                }

                // leave date
                try
                {
                    dimissionDate = Convert.ToDateTime(personInfo.DimissionDate);
                }
                catch (FormatException)
                {
                    dimissionDate = DateTime.MinValue;
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message);
                    return;
                }

                if ((onboardDate != DateTime.MinValue) 
                    && (nextInfo.Date < onboardDate)
                    && ((nextInfo.State == AttendanceState.None) || (nextInfo.State == AttendanceState.Late)))
                {
                    nextInfo.State = AttendanceState.Absent;
                }

                if ((dimissionDate != DateTime.MinValue) 
                    && (nextInfo.Date > dimissionDate)
                    && ((nextInfo.State == AttendanceState.None) || (nextInfo.State == AttendanceState.Late)))
                {
                    nextInfo.State = AttendanceState.Absent;
                }

            } // foreach (var nextInfo in attendanceList)
            #endregion // befor or later than enter and leave date
        }
    }
}
