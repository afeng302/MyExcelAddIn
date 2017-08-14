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




        //private void btnSearch_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        System.Data.DataTable dt = new System.Data.DataTable();

        //        String ConnectionString = "Data Source=YOURDATASOURCE;Initial Catalog=YOURDATABASENAME;User id = UID;password=password";
        //        SqlConnection con = new SqlConnection(ConnectionString);
        //        String Query = " Select Item_Code,Item_Name FROM ItemMasters Where Item_Name LIKE '" + txtItemName.Text.Trim() + "%'";
        //        SqlCommand cmd = new SqlCommand(Query, con);
        //        cmd.CommandType = System.Data.CommandType.Text;
        //        System.Data.SqlClient.SqlDataAdapter sda = new System.Data.SqlClient.SqlDataAdapter(cmd);
        //        sda.Fill(dt);

        //        if (dt.Rows.Count <= 0)
        //        {
        //            return;
        //        }

        //        Globals.ThisAddIn.Application.ActiveSheet.Cells.ClearContents();

        //        Globals.ThisAddIn.Application.ActiveSheet.Cells[1, 1].Value2 = "Item Code";

        //        Globals.ThisAddIn.Application.ActiveSheet.Cells[1, 2].Value2 = "Item Name";

        //        for (int i = 0; i <= dt.Rows.Count - 1; i++)
        //        {

        //            Globals.ThisAddIn.Application.ActiveSheet.Cells[i + 2, 1].Value2 = dt.Rows[i][0].ToString();


        //            Globals.ThisAddIn.Application.ActiveSheet.Cells[i + 2, 2].Value2 = dt.Rows[i][1].ToString();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //}

        private void btnAddText_Click(object sender, EventArgs e)
        {
            // Turn off screen updating and displaying alerts
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            Globals.ThisAddIn.Application.AskToUpdateLinks = false;

            Excel.Range objRange = Globals.ThisAddIn.Application.ActiveCell;

            Excel.Worksheet sheet = Globals.ThisAddIn.Application.Worksheets[1];

            this.DrawHeader(sheet);

            // open files
            Excel.Workbook attendanceBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\科技部外包考勤.xlsx");
            List<AttendanceInfo> unsualInfoList = this.GetUnusalAttendance(attendanceBook.Worksheets[1]);

            // close files
            attendanceBook.Close();

            // Person Repository
            Excel.Workbook personBook = Globals.ThisAddIn.Application.Workbooks.Open("C:\\data\\外包人员台账.xlsx");
            PersonInfoRepo.GenerateInfoMap(personBook);
            personBook.Close();

            // write unsual record
            this.WriteUnsualInfo(unsualInfoList, sheet);

            // Turn on screen updating and displaying alerts again
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.DisplayAlerts = true;
            Globals.ThisAddIn.Application.AskToUpdateLinks = true;
        }

        private void btnAddImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "*";
            dlg.DefaultExt = "bmp";
            dlg.ValidateNames = true;

            dlg.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                Bitmap dImg = new Bitmap(dlg.FileName);

                Excel.Shape IamgeAdd = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPicture(dlg.FileName,

        Microsoft.Office.Core.MsoTriState.msoFalse,

        Microsoft.Office.Core.MsoTriState.msoCTrue,

        20, 30, dImg.Width, dImg.Height);
            }

            //we should also clear the clip board

            System.Windows.Forms.Clipboard.Clear();
        }


        private void DrawHeader(Excel.Worksheet sheet)
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

        private void WriteUnsualInfo(List<AttendanceInfo> infoList, Excel.Worksheet sheet)
        {
            int rowIndex = 1;
            int colIndex = 1;

            foreach (var nextInfo in infoList)
            {
                rowIndex++; // from row #2
                colIndex = 1;

                Excel.Range objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "姓名";
                objRange.Value = nextInfo.Name;
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                PersonInfo personInfo = PersonInfoRepo.GetPersonInfo(nextInfo.Name);

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
                //objRange.Value = "日期";
                objRange.Value = nextInfo.ArriveTime.ToShortDateString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "上班打卡时间";
                objRange.Value = nextInfo.ArriveTime.ToShortTimeString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "下班打卡时间";
                objRange.Value = nextInfo.LeaveTime.ToShortTimeString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, colIndex++];
                //objRange.Value = "备注";
                objRange.Value = nextInfo.State.ToString();
                objRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
        }

        private List<AttendanceInfo> GetUnusalAttendance(Excel.Worksheet srcSheet)
        {
            int srcRowIndex = 1;
            int descRowIndex = 2; // the first line is header

            // 姓名 日期  上班打卡时间 下班打卡时间

            AttendanceInfo yesterdayInfo = null;

            List<AttendanceInfo> infoList = new List<AttendanceInfo>();

            for (srcRowIndex = 2; srcRowIndex < 10000; srcRowIndex++)
            {
                // name
                string name = srcSheet.Cells[srcRowIndex, 1].Value;
                if (string.IsNullOrWhiteSpace(name))
                {
                    break;
                }

                AttendanceInfo todayInfo = new AttendanceInfo(srcSheet.Cells[srcRowIndex, 1].Value,
                    srcSheet.Cells[srcRowIndex, 2].Value, srcSheet.Cells[srcRowIndex, 3].Value,
                    srcSheet.Cells[srcRowIndex, 4].Value);

                infoList.Add(todayInfo);
            }

            infoList.Sort((info1, info2) =>
            {
                if (!info1.Name.Equals(info2.Name))
                {
                    return info1.Name.CompareTo(info2.Name);
                }

                return (int)(info1.ArriveTime - info2.ArriveTime).TotalHours;
            });

            foreach (var todayInfo in infoList)
            {
                this.SetState(todayInfo, yesterdayInfo);

                if (todayInfo.State != AttendanceState.None)
                {
                    Trace.WriteLine(string.Format("*** [{0}]  [{1}]  [{2}]", todayInfo.Name, todayInfo.State, todayInfo.ArriveTime));
                }

                yesterdayInfo = todayInfo;
            }

            List<AttendanceInfo> unsualInfo = new List<AttendanceInfo>();
            infoList.ForEach(x
                =>
            {
                if (x.State != AttendanceState.None)
                {
                    unsualInfo.Add(x);
                }
            });

            return unsualInfo;
        }

        private void SetState(AttendanceInfo todayInfo, AttendanceInfo yesterdayInfo)
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

    }
}
