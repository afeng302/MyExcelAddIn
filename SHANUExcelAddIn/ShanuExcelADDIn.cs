using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Diagnostics;
using SHANUExcelAddIn.Util;
using System.IO;
/// <summary>
/// Author      : Shanu
/// Create date : 2015-02-23
/// Description :Excel AddIn Control
/// Latest
/// Modifier    :Shanu
/// Modify date :  2015-02-23
/// </summary>

namespace SHANUExcelAddIn
{
    public partial class ShanuExcelADDIn : UserControl
    {

        public ShanuExcelADDIn()
        {
            InitializeComponent();


        }

        private void btnAttendanceException_Click(object sender, EventArgs e)
        {
            // Turn off screen updating and displaying alerts
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.AskToUpdateLinks = false;

            try
            {
                Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

                activeSheet.Name = "考勤异常";

                this.DrawUnsualHeader(activeSheet);

                // get attendance info
                Excel.Workbook attendanceBook = null;
                if (File.Exists("C:\\data\\科技部外包考勤.xls"))
                {
                    attendanceBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\科技部外包考勤.xls");
                }
                else if (File.Exists("C:\\data\\科技部外包考勤.xlsx"))
                {
                    attendanceBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\科技部外包考勤.xlsx");
                }
                else
                {
                    MessageBox.Show("[科技部外包考勤] 文件不存在");
                    return;
                }

                List<AttendanceInfo> attendanceInfoList = AttendanceUtil.GetAttendanceInfoList(attendanceBook.Worksheets[1]);

                // close files
                attendanceBook.Close();

                // Person Repository
                Excel.Workbook personBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\外包人员台账.xlsx");
                PersonInfoRepo.GenerateInfoMapByName(personBook);
                personBook.Close();

                // get unsual info 
                List<AttendanceInfo> unsualInfoList = AttendanceUtil.GetUnusalAttendance(attendanceInfoList);

                

                // get no show list
                List<PersonInfo> outsourceList = PersonInfoRepo.GetOnsiteOutsourceList();
                List<PersonInfo> noShowList = AttendanceUtil.GetNoShowPersonList(outsourceList, attendanceInfoList);

                // write unsual record
                this.WriteUnsualInfo(unsualInfoList, noShowList, activeSheet);

            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }

            // Turn on screen updating and displaying alerts again
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            Globals.ThisAddIn.Application.AskToUpdateLinks = true;
        }

        private void DrawUnsualHeader(Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "姓名";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属公司";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "项目组";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "主管项目经理";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属中心";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "日期";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "上班打卡时间";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "下班打卡时间";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "备注";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void WriteUnsualInfo(List<AttendanceInfo> attendanceInfoList, List<PersonInfo> noShowList, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            // unsual attendance
            foreach (var nextInfo in attendanceInfoList)
            {
                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                if (personInfo == null)
                {
                    Trace.WriteLine("cannot find " + nextInfo.Name);
                    continue;
                }

                // check if the person has left
                //if (!string.IsNullOrWhiteSpace(personInfo.LeaveDate)
                //    && (personInfo.LeaveDate != "长期"))
                //{
                //    Trace.WriteLine(personInfo.Name + " has left at " + personInfo.LeaveDate);
                //    continue;
                //}

                rowIndex++; // from row #2
                colIndex = 1;

                this.WriteUnsualAttendanceRow(sheet, rowIndex, colIndex, nextInfo, personInfo);
            }

            // no show list
            foreach (var nextInfo in noShowList)
            {
                rowIndex++;
                colIndex = 1;

                AttendanceInfo attendanceInfo = new AttendanceInfo(nextInfo.Name, string.Empty, string.Empty, string.Empty);
                attendanceInfo.State = AttendanceState.NoShow;

                this.WriteUnsualAttendanceRow(sheet, rowIndex, colIndex, attendanceInfo, nextInfo);
            }
        }

        private void WriteUnsualAttendanceRow(Excel.Worksheet sheet, int rowIndex, int colIndex,
            AttendanceInfo attendanceInfo, PersonInfo personInfo)
        {
            Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "姓名";
            objRange.Value = attendanceInfo.Name;
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "所属公司";
            objRange.Value = personInfo != null ? personInfo.Company : string.Empty;
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "项目组";
            objRange.Value = personInfo != null ? personInfo.Project : string.Empty;
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "主管项目经理";
            objRange.Value = personInfo != null ? personInfo.Manager : string.Empty;
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "所属中心";
            objRange.Value = personInfo != null ? personInfo.Department : string.Empty;
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "日期";
            objRange.Value = attendanceInfo.Date.ToShortDateString();
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "上班打卡时间";
            if (attendanceInfo.State == AttendanceState.Late)
            {
                objRange.Value = attendanceInfo.ArriveTime.ToShortTimeString();
            }
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "下班打卡时间";
            if (attendanceInfo.State == AttendanceState.Late)
            {
                objRange.Value = attendanceInfo.LeaveTime.ToShortTimeString();
            }
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            //objRange.Value = "备注";
            switch (attendanceInfo.State)
            {
                case AttendanceState.Late:
                    objRange.Value = "迟到";
                    break;
                case AttendanceState.Absent:
                    objRange.Value = "缺勤";
                    break;
                case AttendanceState.Left:
                    if (string.IsNullOrWhiteSpace(personInfo.DimissionDate))
                    {
                        objRange.Value = "缺勤-已离场？"; // did not update the status
                    }
                    else
                    {
                        objRange.Value = "缺勤";
                    }
                    break;
                case AttendanceState.NoShow:
                    objRange.Value = "无考勤记录？";
                    break;
                default:
                    break;
            }
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void btnStaffStatistic_Click(object sender, EventArgs e)
        {
            // Turn off screen updating and displaying alerts
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.AskToUpdateLinks = false;

            try
            {
                // Person Repository
                Excel.Workbook personBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\外包人员台账.xlsx");
                PersonInfoRepo.GenerateInfoMapByName(personBook);
                personBook.Close();

                // filter out person
                List<PersonInfo> outsourceList = PersonInfoRepo.GetOnsiteOutsourceList();

                // 
                this.WriteOutsourceInfo(outsourceList, Globals.ThisAddIn.Application.ActiveSheet);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());

                throw;
            }

            // Turn on screen updating and displaying alerts again
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            Globals.ThisAddIn.Application.AskToUpdateLinks = true;
        }

        private void WriteOutsourceInfo(List<PersonInfo> infoList, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "姓名";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属公司";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "项目组";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属系统";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "主管项目经理";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "外包形式";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属中心";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "备注";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            foreach (var nextInfo in infoList)
            {
                rowIndex++; // from row #2
                colIndex = 1;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "姓名";
                objRange.Value = nextInfo.Name;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属公司";
                objRange.Value = nextInfo != null ? nextInfo.Company : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "项目组";
                objRange.Value = nextInfo != null ? nextInfo.Project : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属系统";
                objRange.Value = nextInfo != null ? nextInfo.System : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "主管项目经理";
                objRange.Value = nextInfo != null ? nextInfo.Manager : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "外包形式";
                objRange.Value = nextInfo != null ? nextInfo.WorkType : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属中心";
                objRange.Value = nextInfo != null ? nextInfo.Department : string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "备注";
                objRange.Value = string.Empty;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
        }

        private void btnWorkLoad_Click(object sender, EventArgs e)
        {
            // Turn off screen updating and displaying alerts
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.AskToUpdateLinks = false;

            try
            {
                Excel.Workbook attendanceBook = null;

                // get attendance info
                if (File.Exists("C:\\data\\科技部外包考勤.xls"))
                {
                    attendanceBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\科技部外包考勤.xls");
                }
                else if (File.Exists("C:\\data\\科技部外包考勤.xlsx"))
                {
                    attendanceBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\科技部外包考勤.xlsx");
                }
                else
                {
                    MessageBox.Show("[科技部外包考勤] 文件不存在");
                    return;
                }

                List<AttendanceInfo> attendanceInfoList = AttendanceUtil.GetAttendanceInfoList(attendanceBook.Worksheets[1]);
                attendanceBook.Close();

                // Person Repository
                Excel.Workbook personBook = null;
                if (File.Exists("C:\\data\\外包人员台账.xls"))
                {
                    personBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\外包人员台账.xls");
                }
                else if (File.Exists("C:\\data\\外包人员台账.xlsx"))
                {
                    personBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\外包人员台账.xlsx");
                }
                else
                {
                    MessageBox.Show("[外包人员台账] 文件不存在");
                    return;
                }
                PersonInfoRepo.GenerateInfoMapByName(personBook);
                personBook.Close();

                // get unsual info
                // invoke this method to set the attendance state
                List<AttendanceInfo> unsualInfoList = AttendanceUtil.GetUnusalAttendance(attendanceInfoList);

                // get workload list
                List<WorkloadInfo> workloadListPerMonth = WorkloadUtil.GetWorklaodListPerMonth(attendanceInfoList);

                // write to sheet - per month
                object sheet = Globals.ThisAddIn.Application.Worksheets.Add();
                Globals.ThisAddIn.Application.ActiveSheet.Name = "月度统计";
                this.WriteWorkLoadPerMonth(workloadListPerMonth, Globals.ThisAddIn.Application.ActiveSheet);

                // write to sheet - total
                sheet = Globals.ThisAddIn.Application.Worksheets.Add();
                Globals.ThisAddIn.Application.ActiveSheet.Name = "汇总统计";
                List<WorkloadInfo> workloadListTotally = WorkloadUtil.GetWorkloadListTotally(workloadListPerMonth);
                this.WriteWorkLoadTotally(workloadListTotally, Globals.ThisAddIn.Application.ActiveSheet);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());

                throw;
            }

            // Turn on screen updating and displaying alerts again
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            Globals.ThisAddIn.Application.AskToUpdateLinks = true;
        }

        private void WriteWorkLoadPerMonth(List<WorkloadInfo> workloadList, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            #region Header
            Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "姓名";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属公司";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "项目组";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属系统";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "主管项目经理";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属中心";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "月份";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "应出勤天数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "实出勤天数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "结算人月";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //sheet.Columns[colIndex - 1].Numberformat = "@";
            //sheet.Columns[colIndex - 1].Numberformat = "0.00";

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "迟到天数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "加班小时数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "备注";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            #endregion // Header

            #region rows
            foreach (var nextInfo in workloadList)
            {
                rowIndex++; // from row #2
                colIndex = 1;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "姓名";
                objRange.Value = nextInfo.Name;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                if (personInfo == null)
                {
                    //Debug.Assert(false, nextInfo.Name + " dos not exist");
                    continue;
                }

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属公司";
                objRange.Value = personInfo.Company;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "项目组";
                objRange.Value = personInfo.Project;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属系统";
                objRange.Value = personInfo.System;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "主管项目经理";
                objRange.Value = personInfo.Manager;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属中心";
                objRange.Value = personInfo.Department;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "月份";
                objRange.Value = nextInfo.Month.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "应出勤天数";
                objRange.Value = nextInfo.DueShowDays.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "实出勤天数";
                objRange.Value = nextInfo.ActualShowDays.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "结算人月";
                objRange.Value = string.Format("{0:0.00}", nextInfo.PayStaffMonth);
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "迟到/早退天数";
                objRange.Value = nextInfo.LateDays.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "加班小时数";
                objRange.Value = nextInfo.OTHours.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "备注";
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            } // foreach (var nextInfo in workloadList)
            #endregion // rows
        }

        private void WriteWorkLoadTotally(List<WorkloadInfo> workloadList, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            #region Header
            Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "姓名";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属公司";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "项目组";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属系统";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "主管项目经理";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "所属中心";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "结算人月";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //sheet.Columns[colIndex - 1].Numberformat = "@";

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "迟到折算旷工天数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "加班小时数";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[rowIndex, colIndex++];
            objRange.Value = "备注";
            objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            #endregion // Header

            #region rows
            foreach (var nextInfo in workloadList)
            {
                rowIndex++; // from row #2
                colIndex = 1;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "姓名";
                objRange.Value = nextInfo.Name;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);
                if (personInfo == null)
                {
                    //Debug.Assert(false, nextInfo.Name + " dos not exist");
                    continue;
                }

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属公司";
                objRange.Value = personInfo.Company;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "项目组";
                objRange.Value = personInfo.Project;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属系统";
                objRange.Value = personInfo.System;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "主管项目经理";
                objRange.Value = personInfo.Manager;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "所属中心";
                objRange.Value = personInfo.Department;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "结算人月";
                objRange.Value = string.Format("{0:0.00}", nextInfo.PayStaffMonth);
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "迟到折算旷工天数";
                objRange.Value = nextInfo.LateDays.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "加班小时数";
                objRange.Value = nextInfo.OTHours.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "备注";
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            } // foreach (var nextInfo in workloadList)
            #endregion // rows
        }
    }
}
