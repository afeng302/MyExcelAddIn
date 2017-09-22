using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace SHANUExcelAddIn
{
    static class PersonInfoRepo
    {
        static Dictionary<string, PersonInfo> InfoMap = new Dictionary<string, PersonInfo>();


        public static PersonInfo GetPersonInfo(string name)
        {
            PersonInfo info = null;

            if (InfoMap.TryGetValue(name, out info))
            {
                return info;
            }

            return null;
        }

        public static void GenerateInfoMapByName(Excel.Workbook book)
        {
            int rowIndex = 1;

            int sheetIndex = 1;
            Excel.Worksheet sheet = book.Sheets[sheetIndex];
            while ((sheet.Name != "外包员工资料") && (sheetIndex++ < book.Sheets.Count))
            {
                sheet = book.Sheets[sheetIndex];
            }


            for (rowIndex = 1; rowIndex < 5000; rowIndex++)
            {
                PersonInfo info = new PersonInfo();

                // name
                info.Name = sheet.Cells[rowIndex, 3].Value;
                if (string.IsNullOrWhiteSpace(info.Name))
                {
                    continue;
                }

                // company
                info.Company = sheet.Cells[rowIndex, 2].Value;

                // manager
                info.Manager = sheet.Cells[rowIndex, 4].Value;

                // project
                info.Project = sheet.Cells[rowIndex, 5].Value;

                // work type
                info.WorkType = sheet.Cells[rowIndex, 6].Value;

                // department
                info.Department = sheet.Cells[rowIndex, 8].Value;

                // enter date
                info.EnterDate = sheet.Cells[rowIndex, 10].Value != null ? sheet.Cells[rowIndex, 10].Value.ToString() : null;

                // leave date
                info.LeaveDate = sheet.Cells[rowIndex, 11].Value != null ? sheet.Cells[rowIndex, 11].Value.ToString() : null;


                lock (InfoMap)
                {
                    InfoMap[info.Name] = info;
                }
            } // for (rowIndex = 1; rowIndex < 5000; rowIndex++)

            // correct data and set owner system
            CorrectData(InfoMap.Values);
        }

        public static List<PersonInfo> GetOnsiteOutsourceList()
        {
            List<PersonInfo> outsourceList = new List<PersonInfo>();

            // filter out 
            foreach (var nextPerson in InfoMap.Values)
            {
                if (nextPerson.WorkType != "人力")
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(nextPerson.LeaveDate) && (nextPerson.LeaveDate != "长期"))
                {
                    continue;
                }

                outsourceList.Add(nextPerson);
            }

            // sort
            outsourceList.Sort((info1, info2) =>
            {
                if (info1.Project != info2.Project)
                {
                    return info1.Project != null ? info1.Project.CompareTo(info2.Project) : 1;
                }

                return info1.Company != null ? info1.Company.CompareTo(info2.Company) : 1;
            });

            return outsourceList;
        }

        static void CorrectData(ICollection<PersonInfo> personList)
        {
            #region correct department
            foreach (var nextPerson in personList)
            {
                if (nextPerson.Company != null && nextPerson.Company.Contains("捷科"))
                {
                    nextPerson.Department = "测试中心";
                    continue;
                }

                if (nextPerson.Company != null && nextPerson.Company.Contains("江融信"))
                {
                    nextPerson.Department = "开发中心";
                    continue;
                }

                if (nextPerson.Project != null && nextPerson.Project.Contains("大数据"))
                {
                    nextPerson.Department = "大数据中心";
                    continue;
                }

                if (nextPerson.Manager != null && nextPerson.Manager.Contains("张宪杰"))
                {
                    nextPerson.Department = "大数据中心";
                    continue;
                }
            }
            #endregion //correct department

            #region set owner system
            foreach (var nextPerson in personList)
            {
                if (string.IsNullOrWhiteSpace(nextPerson.Project))
                {
                    Trace.WriteLine("nextInfo.Project is empty");
                    continue;
                }

                // correct "消费信贷（一期）"
                if (nextPerson.Project.Contains("消费信贷（一期）")
                    || nextPerson.Project.Contains("消费信贷（二期）"))
                {
                    nextPerson.Project = "消费信贷";
                }

                // 个人信贷系统
                if (nextPerson.Project.Contains("消费信贷")
                || nextPerson.Project.Contains("操作平台")
                || nextPerson.Project.Contains("openapi")
                || nextPerson.Project.Contains("信贷核心")
                || nextPerson.Project.Contains("资信平台")
                || nextPerson.Project.Contains("调度平台"))
                {
                    nextPerson.System = "个人信贷系统";
                    continue;
                }

                // 统一支付平台
                if (nextPerson.Project.Contains("统一支付"))
                {
                    nextPerson.System = "统一支付平台";
                    continue;
                }

                // 个人理财系统
                if (nextPerson.Project.Contains("个人理财"))
                {
                    nextPerson.System = "个人理财系统";
                    continue;
                }

                // 开放平台
                if (nextPerson.Project.Contains("开放平台"))
                {
                    nextPerson.System = "开放平台";
                    continue;
                }

                // 互金平台
                if (nextPerson.Project.Contains("互联网金融平台")
                    || nextPerson.Project.Contains("互金平台"))
                {
                    nextPerson.System = "互金平台";
                    continue;
                }

                // 对公信贷系统
                if (nextPerson.Project.Contains("对公信贷"))
                {
                    nextPerson.System = "对公信贷系统";
                    continue;
                }

                // 客服系统
                if (nextPerson.Project.Contains("客服"))
                {
                    nextPerson.System = "客服系统";
                    continue;
                }

                // 渠道类系统（APP，微信）
                if (nextPerson.Project.Contains("APP")
                    || nextPerson.Project.Contains("渠道")
                    || nextPerson.Project.Contains("H5")
                    || (nextPerson.Project.Contains("微信") && (nextPerson.Manager == "王月超")))
                {
                    nextPerson.System = "渠道类系统";
                    continue;
                }

                // 财管系统
                if (nextPerson.Project.Contains("财管")
                    || nextPerson.Project.Contains("总账"))
                {
                    nextPerson.System = "财管系统";
                    continue;
                }

                // 促销系统
                if (nextPerson.Project.Contains("促销"))
                {
                    nextPerson.System = "促销系统";
                    continue;
                }

                // 大核心银行（含联网核查，电信反诈骗等）
                if (nextPerson.Project.Contains("核心银行")
                    || nextPerson.Project.Contains("核心系统")
                    || nextPerson.Project.Contains("联网核查")
                    || nextPerson.Project.Contains("电信反诈骗"))
                {
                    nextPerson.System = "大核心银行（含联网核查，电信反诈骗等）";
                    continue;
                }
            }
            #endregion // set owner system
        }
    }
}
