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

            }
        }

        public static List<PersonInfo> GetOutsourceList()
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

            // correct department
            foreach (var nextInfo in outsourceList)
            {
                if (nextInfo.Company != null && nextInfo.Company.Contains("捷科"))
                {
                    nextInfo.Department = "测试中心";
                    continue;
                }

                if (nextInfo.Company != null && nextInfo.Company.Contains("江融信"))
                {
                    nextInfo.Department = "开发中心";
                    continue;
                }

                if (nextInfo.Project!= null && nextInfo.Project.Contains("大数据"))
                {
                    nextInfo.Department = "大数据中心";
                    continue;
                }

                if (nextInfo.Manager != null && nextInfo.Manager.Contains("张宪杰"))
                {
                    nextInfo.Department = "大数据中心";
                    continue;
                }
            }

            // set owner system
            foreach (var nextInfo in outsourceList)
            {
                if (string.IsNullOrWhiteSpace(nextInfo.Project))
                {
                    Trace.WriteLine("nextInfo.Project is empty");
                    continue;
                }

                // 个人信贷系统
                if (nextInfo.Project.Contains("消费信贷") 
                    || nextInfo.Project.Contains("操作平台")
                    || nextInfo.Project.Contains("openapi")
                    || nextInfo.Project.Contains("信贷核心")
                    || nextInfo.Project.Contains("资信平台")
                    || nextInfo.Project.Contains("调度平台"))
                {
                    nextInfo.System = "个人信贷系统";
                    continue;
                }

                // 统一支付平台
                if (nextInfo.Project.Contains("统一支付"))
                {
                    nextInfo.System = "统一支付平台";
                    continue;
                }

                // 个人理财系统
                if (nextInfo.Project.Contains("个人理财"))
                {
                    nextInfo.System = "个人理财系统";
                    continue;
                }

                // 开放平台
                if (nextInfo.Project.Contains("开放平台"))
                {
                    nextInfo.System = "开放平台";
                    continue;
                }

                // 互金平台
                if (nextInfo.Project.Contains("互联网金融平台")
                    || nextInfo.Project.Contains("互金平台"))
                {
                    nextInfo.System = "互金平台";
                    continue;
                }

                // 对公信贷系统
                if (nextInfo.Project.Contains("对公信贷"))
                {
                    nextInfo.System = "对公信贷系统";
                    continue;
                }

                // 客服系统
                if (nextInfo.Project.Contains("客服"))
                {
                    nextInfo.System = "客服系统";
                    continue;
                }

                // 渠道类系统（APP，微信）
                if (nextInfo.Project.Contains("APP")
                    || nextInfo.Project.Contains("渠道")
                    || nextInfo.Project.Contains("H5")
                    || (nextInfo.Project.Contains("微信") && (nextInfo.Manager == "王月超")))
                {
                    nextInfo.System = "渠道类系统";
                    continue;
                }

                // 财管系统
                if (nextInfo.Project.Contains("财管")
                    || nextInfo.Project.Contains("总账"))
                {
                    nextInfo.System = "财管系统";
                    continue;
                }

                // 促销系统
                if (nextInfo.Project.Contains("促销"))
                {
                    nextInfo.System = "促销系统";
                    continue;
                }

                // 大核心银行（含联网核查，电信反诈骗等）
                if (nextInfo.Project.Contains("核心银行")
                    || nextInfo.Project.Contains("核心系统")
                    || nextInfo.Project.Contains("联网核查")
                    || nextInfo.Project.Contains("电信反诈骗"))
                {
                    nextInfo.System = "大核心银行（含联网核查，电信反诈骗等）";
                    continue;
                }

            } // foreach (var nextInfo in outsourceList)

            return outsourceList;
        }
    }
}
