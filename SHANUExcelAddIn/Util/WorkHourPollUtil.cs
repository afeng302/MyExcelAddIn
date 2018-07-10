using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SHANUExcelAddIn.Util
{
    class WorkHourItem
    {
        public string Name { get; set; }

        public DateTime EnterDate { get; set; }

        public DateTime LeaveDate { get; set; }

        public DateTime Month { get; set; }

        public string Department { get; set; }

        public string System { get; set; }

        public double StaffMonth { get; set; }

        public WorkHourItem Clone()
        {
            WorkHourItem other = new WorkHourItem();

            other.Name = this.Name;
            other.EnterDate = this.EnterDate;
            other.LeaveDate = this.LeaveDate;
            other.Month = this.Month;
            other.Department = this.Department;
            other.System = this.System;
            other.StaffMonth = this.StaffMonth;

            return other;
        }
    }

    class WorkHourPollUtil
    {
        static List<WorkHourItem> WORK_HOUR_ITEM_LIST = new List<WorkHourItem>();

        public static void CalcStaffMonth(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            WORK_HOUR_ITEM_LIST.Clear();

            int nameColumnIndex = 0;
            int enterColumnIndex = 0;
            int leaveColumnIndex = 0;
            int departmentColumnIndex = 0;
            int monthStartColumnIndex = 0;

            DateTime[] months = new DateTime[12 * 3];

            #region locate column index
            for (int i = 1; i < 10; i++)
            {
                string cellValue = sheet.Cells[1, i].Value != null ? sheet.Cells[1, i].Value.ToString() : string.Empty;
                if (cellValue == "员工姓名")
                {
                    nameColumnIndex = i;
                }
                else if (cellValue == "入场日期")
                {
                    enterColumnIndex = i;
                }
                else if (cellValue == "离场日期")
                {
                    leaveColumnIndex = i;
                }
                else if (cellValue == "所属中心")
                {
                    departmentColumnIndex = i;
                }
                else if (monthStartColumnIndex == 0)
                {
                    try
                    {
                        Convert.ToDateTime(cellValue);
                        monthStartColumnIndex = i;

                        for (int j = 0; j < 12 * 3; j++)
                        {
                            int validColNum = monthStartColumnIndex + j - j % 3;
                            months[j] = Convert.ToDateTime(sheet.Cells[1, validColNum].Value);
                        }
                    }
                    catch (Exception)
                    {
                        Trace.WriteLine("not valid date format: " + cellValue);
                    }
                }

            } // for (int i = 1; i < 10; i++)

            #endregion // locate column index

            #region read date from sheet
            for (int rowIndex = 2; rowIndex < 500; rowIndex++)
            {
                WorkHourItem masterItem = new WorkHourItem();

                // name
                masterItem.Name = sheet.Cells[rowIndex, nameColumnIndex].Value;
                if (string.IsNullOrWhiteSpace(masterItem.Name))
                {
                    break; // reach the end of list
                }

                // enter date
                try
                {
                    masterItem.EnterDate = Convert.ToDateTime(sheet.Cells[rowIndex, enterColumnIndex].Value);
                }
                catch (Exception)
                {
                    MessageBox.Show("invalid date format. " + masterItem.Name + " " + sheet.Cells[rowIndex, enterColumnIndex].Value);
                }

                // leave date
                try
                {
                    masterItem.LeaveDate = Convert.ToDateTime(sheet.Cells[rowIndex, leaveColumnIndex].Value);
                }
                catch (Exception)
                {
                    MessageBox.Show("invalid date format. " + masterItem.Name + " " + sheet.Cells[rowIndex, leaveColumnIndex].Value);
                }

                // department
                masterItem.Department = sheet.Cells[rowIndex, departmentColumnIndex].Value;

                // system name
                for (int colIndex = monthStartColumnIndex; colIndex < monthStartColumnIndex + 12 * 3; colIndex++)
                {
                    WorkHourItem item = masterItem.Clone();
                    item.System = sheet.Cells[rowIndex, colIndex].Value;
                    if (string.IsNullOrWhiteSpace(item.System))
                    {
                        continue;
                    }

                    item.Month = months[colIndex - monthStartColumnIndex];
                    WORK_HOUR_ITEM_LIST.Add(item);
                }

            } // for (int i = 0; i < 500; i++)
            #endregion // read date from sheet

            Dictionary<string, List<WorkHourItem>> nameMonthMap = new Dictionary<string, List<WorkHourItem>>();
            Dictionary<string, WorkHourItem> systemMap = new Dictionary<string, WorkHourItem>();
            List<WorkHourItem> removalItemList = new List<WorkHourItem>();

            #region calculate staff month

            //
            // build the name-month map
            foreach (var nextItem in WORK_HOUR_ITEM_LIST)
            {
                string key = string.Format("{0}.{1}", nextItem.Name, nextItem.Month);
                if (!nameMonthMap.ContainsKey(key))
                {
                    nameMonthMap[key] = new List<WorkHourItem>();
                }
                nameMonthMap[key].Add(nextItem);
            }

            //
            // calculate the work hours for each system
            foreach (var nextKey in nameMonthMap.Keys)
            {
                double averageStaffMonth = Math.Round((double)1 / nameMonthMap[nextKey].Count, 2);

                systemMap.Clear();

                for (int i = 0; i < nameMonthMap[nextKey].Count; i++)
                {
                    WorkHourItem nextItem = nameMonthMap[nextKey][i];

                    // calculate the staff month
                    if (i == nameMonthMap[nextKey].Count - 1) // the last one
                    {
                        nextItem.StaffMonth = Math.Round(1 - averageStaffMonth * (nameMonthMap[nextKey].Count - 1), 2);
                    }
                    else
                    {
                        nextItem.StaffMonth = averageStaffMonth;
                    }

                    // check duplicate system
                    if (systemMap.ContainsKey(nextItem.System))
                    {
                        systemMap[nextItem.System].StaffMonth += nextItem.StaffMonth;
                        removalItemList.Add(nextItem);
                    }
                    else
                    {
                        systemMap[nextItem.System] = nextItem;
                    }
                } // for (int i = 0; i < nameMonthMap[nextKey].Count; i++)
            } // foreach (var nextKey in nameMonthMap.Keys)

            //
            // remove the duplicat items
            foreach (var nextItem in removalItemList)
            {
                WORK_HOUR_ITEM_LIST.Remove(nextItem);
            }

            // 
            // sort by system, name, month
            WORK_HOUR_ITEM_LIST.Sort((x, y) =>
            {
                if (x.System != y.System)
                {
                    return x.System.CompareTo(y.System);
                }
                if (x.Name != y.Name)
                {
                    return x.Name.CompareTo(y.Name);
                }
                return x.Month.CompareTo(y.Month);
            });
            #endregion // calculate staff month
        }

        public static void WriteLines(Microsoft.Office.Interop.Excel.Worksheet sheet)
        {
            int colIndex = 0;
            int rowIndex = 0;
            Microsoft.Office.Interop.Excel.Range objRange = null;

            #region title line
            rowIndex = 1;
            colIndex = 1;
            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "参与项目";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "员工姓名";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "入场日期";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "离场日期";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            for (int i = 0; i < 12; i++)
            {
                objRange = sheet.Cells[rowIndex, colIndex++];
                objRange.Value = string.Format("{0}月", i + 1);
                objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            }
            #endregion title line

            Dictionary<string, List<WorkHourItem>> systemNameMap = new Dictionary<string, List<WorkHourItem>>();

            #region Group by system + name
            foreach (var nextItem in WORK_HOUR_ITEM_LIST)
            {
                string key = string.Format("{0}.{1}", nextItem.System, nextItem.Name);
                if (!systemNameMap.ContainsKey(key))
                {
                    systemNameMap[key] = new List<WorkHourItem>();
                }
                systemNameMap[key].Add(nextItem);
            }
            #endregion Group by system + name

            #region write out lines
            foreach (var nextKey in systemNameMap.Keys)
            {
                foreach (var nextItem in systemNameMap[nextKey])
                {
                    // write header
                    if (nextItem == systemNameMap[nextKey].First<WorkHourItem>())
                    {
                        // move the next line
                        rowIndex++;
                        colIndex = 1;

                        objRange = sheet.Cells[rowIndex, colIndex++];
                        objRange.Value = nextItem.System;
                        objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        objRange = sheet.Cells[rowIndex, colIndex++];
                        objRange.Value = nextItem.Name;
                        objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        objRange = sheet.Cells[rowIndex, colIndex++];
                        objRange.Value = nextItem.EnterDate;
                        objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        objRange = sheet.Cells[rowIndex, colIndex++];
                        objRange.Value = nextItem.LeaveDate;
                        objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    }

                    // write effort for each month
                    int startColIndex = 5; // the Janary column number (1 base)
                    colIndex = startColIndex + nextItem.Month.Month - 1;
                    objRange = sheet.Cells[rowIndex, colIndex];
                    objRange.Value = nextItem.StaffMonth;
                    objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                } // foreach (var nextItem in systemNameMap[nextKey])
            } // foreach (var nextKey in systemNameMap.Keys)
            #endregion

        }

    }
}
