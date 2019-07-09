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
            int nameColumnIndex = 0;
            int dateColumnIndex = 0;
            int arriveColumnIndex = 0;
            int leaveColumnIndex = 0;
            int exceptionColumnIndex = 0;

            // locate column index
            for (int i = 1; i < 30; i++)
            {
                string cellValue = srcSheet.Cells[1, i].Value != null ? srcSheet.Cells[1, i].Value.ToString() : string.Empty;
                if (cellValue == "姓名")
                {
                    nameColumnIndex = i;
                }
                else if (cellValue == "日期")
                {
                    dateColumnIndex = i;
                }
                else if (cellValue == "签到时间")
                {
                    arriveColumnIndex = i;
                }
                else if (cellValue == "签退时间")
                {
                    leaveColumnIndex = i;
                }
                else if (cellValue == "例外情况")
                {
                    exceptionColumnIndex = i;
                }
            }

            // 姓名 日期  签到时间 签退时间
            List<AttendanceInfo> infoList = new List<AttendanceInfo>();

            for (int startRowIndex = 2; startRowIndex < 50000; startRowIndex++)
            {
                // name
                string name = srcSheet.Cells[startRowIndex, nameColumnIndex].Value;
                if (string.IsNullOrWhiteSpace(name))
                {
                    if (startRowIndex > 10)
                    {
                        break;
                    }
                    continue; // the top lines may have some statistic infomation
                }

                AttendanceInfo todayInfo = new AttendanceInfo(name,
                    srcSheet.Cells[startRowIndex, dateColumnIndex].Value,
                    srcSheet.Cells[startRowIndex, arriveColumnIndex].Value,
                    srcSheet.Cells[startRowIndex, leaveColumnIndex].Value,
                    srcSheet.Cells[startRowIndex, exceptionColumnIndex].Value);

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

            // set attendance state
            foreach (var todayInfo in attendanceList)
            {
                SetAttendanceState(todayInfo, yesterdayInfo);

                if (todayInfo.State != AttendanceState.Normal)
                {
                    Trace.WriteLine(string.Format("*** [{0}]  [{1}]  [{2}]", todayInfo.Name, todayInfo.State, todayInfo.ArriveTime));
                }

                yesterdayInfo = todayInfo;
            }

            // correct the attendance mistake caused by system problem
            AttendanceCorrection(attendanceList);

            // collect the unsual record
            List<AttendanceInfo> unsualInfoList = new List<AttendanceInfo>();
            foreach (var nextInfo in attendanceList)
            {
                if ((nextInfo.State != AttendanceState.Normal) && (nextInfo.State != AttendanceState.PaidLeave))
                {
                    unsualInfoList.Add(nextInfo);
                }
            };

            // set the dismissed case
            SetDismission4AllPerson(unsualInfoList);

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

        public static void FilteroutDissmissedPerson(List<AttendanceInfo> infoList)
        {
            // sort by date to get the first and last day
            infoList.Sort((info1, info2) => info1.Date.CompareTo(info2.Date));
            DateTime firstStatisticDay = infoList.First<AttendanceInfo>().Date;
            DateTime lastStatisticDay = infoList.Last<AttendanceInfo>().Date;

            // get dissmissed person set
            HashSet<string> dissmissedPersonSet = new HashSet<string>();
            foreach (var nextInfo in infoList)
            {
                if (!string.IsNullOrEmpty(nextInfo.Name) && !dissmissedPersonSet.Contains(nextInfo.Name))
                {
                    dissmissedPersonSet.Add(nextInfo.Name);
                }
            }
            foreach (var nextPerson in dissmissedPersonSet.ToList())
            {
                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextPerson);
                if (personInfo == null)
                {
                    MessageBox.Show(string.Format("警告：没有找到台账信息 [{0}] ", nextPerson));
                    continue;
                }

                // check dismission date
                DateTime dimissionDate = DateTime.MaxValue;
                try
                {
                    if (!string.IsNullOrWhiteSpace(personInfo.DimissionDate))
                    {
                        dimissionDate = Convert.ToDateTime(personInfo.DimissionDate);
                    }
                }
                catch (Exception)
                {
                    Trace.WriteLine("error dismission date format: " + personInfo.DimissionDate);
                    dimissionDate = DateTime.MinValue;
                }
                if (firstStatisticDay <= dimissionDate)
                {
                    dissmissedPersonSet.Remove(nextPerson);
                }
                else
                {
                    Trace.WriteLine("dismissed person: " + nextPerson);
                }
            } // foreach (var nextPerson in dissmissedPersonSet.ToList())

            // remove the dissmissed records
            List<AttendanceInfo> removedInfoList = new List<AttendanceInfo>();
            foreach (var nextInfo in infoList)
            {
                if (dissmissedPersonSet.Contains(nextInfo.Name))
                {
                    removedInfoList.Add(nextInfo);
                }
            }
            foreach (var nextInfo in removedInfoList)
            {
                infoList.Remove(nextInfo);
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
            //int averageDays = totaldays / nameMap.Count;
            //foreach (var nextName in nameMap.Keys)
            //{
            //    Debug.Assert(nameMap[nextName].Count <= averageDays, "there is duplicate names?  " + nextName);
            //}

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

            // for pan
            foreach (var nextInfo in infoList)
            {
                if ((nextInfo.Name == "潘俐") && ((nextInfo.Date >= DateTime.Parse("2019-5-5")) && (nextInfo.Date <= DateTime.Parse("2019-5-11"))))
                {
                    nextInfo.ArriveTime = new DateTime(nextInfo.Date.Year, nextInfo.Date.Month, nextInfo.Date.Day, 9, 1, 2);
                    nextInfo.LeaveTime = new DateTime(nextInfo.Date.Year, nextInfo.Date.Month, nextInfo.Date.Day, 19, 12, 12);
                    nextInfo.State = AttendanceState.Normal;
                }

                if ((nextInfo.Name == "潘俐") && ((nextInfo.Date >= DateTime.Parse("2019-5-13")) && (nextInfo.Date <= DateTime.Parse("2019-5-17"))))
                {
                    nextInfo.ArriveTime = new DateTime(nextInfo.Date.Year, nextInfo.Date.Month, nextInfo.Date.Day, 9, 1, 2);
                    nextInfo.LeaveTime = new DateTime(nextInfo.Date.Year, nextInfo.Date.Month, nextInfo.Date.Day, 19, 12, 12);
                    nextInfo.State = AttendanceState.Normal;
                }
            }
        }

        static void SetAttendanceState(AttendanceInfo todayInfo, AttendanceInfo yesterdayInfo)
        {
            // skip the dayoff (weekend and holiday)
            if (!WorkdayUtil.IsWorkday(todayInfo.Date))
            {
                todayInfo.State = AttendanceState.PaidLeave;
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
            if ((yesterdayInfo != null) && (todayInfo.Date.DayOfYear - yesterdayInfo.Date.DayOfYear != 1))
            {
                Trace.WriteLine("date is not contiguous");
                yesterdayInfo = null;
            }

            // leave (no pay)
            if (!string.IsNullOrWhiteSpace(todayInfo.ExceptionComments)
                && (todayInfo.ExceptionComments == "请假"))
            {
                todayInfo.State = AttendanceState.Leave;
                return;
            }

            // additional record (no fine)
            if (!string.IsNullOrWhiteSpace(todayInfo.ExceptionComments)
                && (todayInfo.ExceptionComments == "补录异常"))
            {
                todayInfo.State = AttendanceState.AdditionalRecord;
                return;
            }

            // addiional record (fined)
            // to do ...

            // absent case
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

            // late case
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

            // the nomral case
            todayInfo.State = AttendanceState.Normal;
        }

        static void SetDismission4AllPerson(List<AttendanceInfo> unsualInfo)
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
                SetDismission4OnePerson(nextInfoList);
            }

        }

        static void SetDismission4OnePerson(List<AttendanceInfo> unsualInfo)
        {
            if (unsualInfo.Count < DismissionChecker.MAX_ABSENT_DAYS)
            {
                return;
            }

            DismissionChecker checker = new DismissionChecker(unsualInfo.First<AttendanceInfo>());

            foreach (var nextInfo in unsualInfo)
            {
                checker.GetNextInfo(nextInfo);

                if (checker.IsStateWillChange)
                {
                    if (checker.NextState == AttendanceState.Dimission)
                    {
                        checker.ChangeState();
                        continue;
                    }

                    if (checker.NextState == AttendanceState.Absent)
                    {
                        SetDismissionNodes(unsualInfo, checker.FirstAbsentNode, checker.PreNode);
                        checker.ChangeState();
                    }
                }
            } // foreach (var nextInfo in unsualInfo)

            if (checker.PreState == AttendanceState.Dimission)
            {
                SetDismissionNodes(unsualInfo, checker.FirstAbsentNode, checker.PreNode);
            }
        }

        static void SetDismissionNodes(List<AttendanceInfo> unsualInfo, AttendanceInfo firstNode, AttendanceInfo currNode)
        {
            foreach (var nextInfo in unsualInfo)
            {
                if (nextInfo.Date.DayOfYear < firstNode.Date.DayOfYear)
                {
                    continue;
                }
                nextInfo.State = AttendanceState.Dimission;

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
            // correct system fault
            foreach (var nextInfo in attendanceList)
            {
                if (CorrectSystemFault(nextInfo, new DateTime(2017, 8, 31), AttendanceState.None, AttendanceState.Normal))
                {
                    continue;
                }
                if (CorrectSystemFault(nextInfo, new DateTime(2017, 12, 28), AttendanceState.None, AttendanceState.Normal))
                {
                    continue;
                }
                if (CorrectSystemFault(nextInfo, new DateTime(2017, 12, 29), AttendanceState.None, AttendanceState.Normal))
                {
                    continue;
                }

                // leave 3 days before spring festival
                if (CorrectSystemFault(nextInfo, new DateTime(2018, 2, 12), AttendanceState.Absent, AttendanceState.Leave))
                {
                    continue;
                }
                if (CorrectSystemFault(nextInfo, new DateTime(2018, 2, 13), AttendanceState.Absent, AttendanceState.Leave))
                {
                    continue;
                }
                if (CorrectSystemFault(nextInfo, new DateTime(2018, 2, 14), AttendanceState.Absent, AttendanceState.Leave))
                {
                    continue;
                }

            } // foreach (var nextInfo in attendanceList)

            //
            // befor or later than enter and leave date
            // always put this rule in the last
            // 
            #region // befor or later than enter and leave date
            foreach (var nextInfo in attendanceList)
            {
                DateTime onboardDate = DateTime.MaxValue;
                DateTime dimissionDate = DateTime.MinValue;

                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                if (personInfo == null)
                {
                    //MessageBox.Show(nextInfo.Name + " 没有台账信息!");
                    Trace.WriteLine("cannot find " + nextInfo.Name);
                    continue;
                }

                // onboard date
                try
                {
                    if (!ONBOARD_FORMAT_EXCEPTION.ContainsKey(personInfo.Name))
                    {
                        onboardDate = Convert.ToDateTime(personInfo.OnboardDate);
                    }
                }
                catch (FormatException)
                {
                    Trace.WriteLine("error onboard date format: " + personInfo.OnboardDate);
                    onboardDate = DateTime.MaxValue;
                    ONBOARD_FORMAT_EXCEPTION[personInfo.Name] = personInfo;
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message);
                    return;
                }

                // leave date
                try
                {
                    if (string.IsNullOrWhiteSpace(personInfo.DimissionDate))
                    {
                        dimissionDate = DateTime.MaxValue; // not dismissed
                    }
                    else if (!DISMISSION_FORMAT_EXCEPTION.ContainsKey(personInfo.Name))
                    {
                        dimissionDate = Convert.ToDateTime(personInfo.DimissionDate);
                    }
                }
                catch (FormatException)
                {
                    Trace.WriteLine("error dismission date format: " + personInfo.DimissionDate);
                    dimissionDate = DateTime.MinValue;
                    DISMISSION_FORMAT_EXCEPTION[personInfo.Name] = personInfo;
                }
                catch (Exception exp)
                {
                    MessageBox.Show(exp.Message);
                    return;
                }

                if ((nextInfo.Date < onboardDate) && (nextInfo.State != AttendanceState.PaidLeave))
                //&& ((nextInfo.State == AttendanceState.Normal) || (nextInfo.State == AttendanceState.Late)))
                {
                    nextInfo.State = AttendanceState.NotOnboard;
                }

                if ((nextInfo.Date > dimissionDate) && (nextInfo.State != AttendanceState.PaidLeave))
                //&& ((nextInfo.State == AttendanceState.Normal) || (nextInfo.State == AttendanceState.Late)))
                {
                    nextInfo.State = AttendanceState.Dimission;
                }

            } // foreach (var nextInfo in attendanceList)
            #endregion // befor or later than enter and leave date
        }

        static Dictionary<string, PersonInfo> ONBOARD_FORMAT_EXCEPTION = new Dictionary<string, PersonInfo>();
        static Dictionary<string, PersonInfo> DISMISSION_FORMAT_EXCEPTION = new Dictionary<string, PersonInfo>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="faultDate"></param>
        /// <returns>
        /// false - did not do anything
        /// true - correct successfully
        /// </returns>
        static bool CorrectSystemFault(AttendanceInfo info, DateTime faultDate, AttendanceState errorState, AttendanceState correctState)
        {
            if (!info.Date.Equals(faultDate))
            {
                return false;
            }

            PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(info.Name);
            if (personInfo == null)
            {
                //MessageBox.Show(nextInfo.Name + " 没有台账信息!");
                Trace.WriteLine("cannot find " + info.Name);
                return false;
            }

            // check dismission date
            DateTime dimissionDate = DateTime.MaxValue;
            try
            {
                if (!string.IsNullOrWhiteSpace(personInfo.DimissionDate)
                    && !DISMISSION_FORMAT_EXCEPTION.ContainsKey(personInfo.Name))
                {
                    dimissionDate = Convert.ToDateTime(personInfo.DimissionDate);
                }
            }
            catch (FormatException)
            {
                Trace.WriteLine("error dismission date format: " + personInfo.DimissionDate);
                dimissionDate = DateTime.MinValue;
                DISMISSION_FORMAT_EXCEPTION[personInfo.Name] = personInfo;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
                return false;
            }
            if (dimissionDate < faultDate)
            {
                return false; // skip the leave early case
            }

            // check onboard date
            DateTime onboardDate = DateTime.MinValue;
            try
            {
                if (!string.IsNullOrWhiteSpace(personInfo.OnboardDate)
                    && !ONBOARD_FORMAT_EXCEPTION.ContainsKey(personInfo.Name))
                {
                    onboardDate = Convert.ToDateTime(personInfo.OnboardDate);
                }
            }
            catch (FormatException)
            {
                Trace.WriteLine("error onboard date format: " + personInfo.OnboardDate);
                onboardDate = DateTime.MaxValue;
                ONBOARD_FORMAT_EXCEPTION[personInfo.Name] = personInfo;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
                return false;
            }
            if (onboardDate > faultDate)
            {
                return false; // skip the onboard later case
            }

            // correct the status
            if (errorState == AttendanceState.None)
            {
                info.State = correctState; // for all the case
            }
            else if (info.State == errorState)
            {
                info.State = correctState; // for specific case
            }

            return true;
        }
    }
}
