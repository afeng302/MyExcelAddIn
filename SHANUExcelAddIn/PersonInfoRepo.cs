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
            while ((!sheet.Name.Contains("资料")) && (sheetIndex++ < book.Sheets.Count))
            {
                sheet = book.Sheets[sheetIndex];
            }

            int nameColumn = 0;
            int companyColumn = 0;
            int managerColumn = 0;
            int costProjColumn = 0;
            int bizProjColumn = 0;
            int workTypeColumn = 0;
            int rankColumn = 0;
            int departmentColumn = 0;
            int enterDateColumn = 0;
            int leaveDateColumn = 0;

            #region locate column index
            for (int i = 1; i < 20; i++)
            {
                string cellValue = sheet.Cells[1, i].Value != null ? sheet.Cells[1, i].Value.ToString() : string.Empty;
                if (cellValue == "姓名")
                {
                    nameColumn = i;
                }
                else if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("供应商名称"))
                {
                    companyColumn = i;
                }
                else if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("项目经理"))
                {
                    managerColumn = i;
                }
                else if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("人力成本归集"))
                {
                    costProjColumn = i;
                }
                else if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("参与项目"))
                {
                    bizProjColumn = i;
                }
                else if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("外包形式"))
                {
                    workTypeColumn = i;
                }
                else if (cellValue == "职级")
                {
                    rankColumn = i;
                }
                else if (cellValue == "所属中心")
                {
                    departmentColumn = i;
                }
                else if (cellValue == "入场时间")
                {
                    enterDateColumn = i;
                }
                else if (cellValue == "离场时间")
                {
                    leaveDateColumn = i;
                }

            } // for (int i = 1; i < 20; i++)
            #endregion locate column index

            //
            // get cell values
            List<PersonInfo> personList = new List<PersonInfo>();
            for (rowIndex = 1; rowIndex < 5000; rowIndex++)
            {
                PersonInfo info = new PersonInfo();

                // name
                info.Name = Convert.ToString(sheet.Cells[rowIndex, nameColumn].Value);
                if (string.IsNullOrWhiteSpace(info.Name))
                {
                    continue;
                }

                // company
                info.Company = Convert.ToString(sheet.Cells[rowIndex, companyColumn].Value);

                // manager
                info.Manager = Convert.ToString(sheet.Cells[rowIndex, managerColumn].Value);

                // cost project
                info.CostProject = Convert.ToString(sheet.Cells[rowIndex, costProjColumn].Value);

                // business project
                info.BizProject = Convert.ToString(sheet.Cells[rowIndex, bizProjColumn].Value);

                // work type
                info.WorkType = Convert.ToString(sheet.Cells[rowIndex, workTypeColumn].Value);

                // Rank - 初、中、高、专家
                info.Rank = Convert.ToString(sheet.Cells[rowIndex, rankColumn].Value);

                // department
                info.Department = Convert.ToString(sheet.Cells[rowIndex, departmentColumn].Value);

                // enter date
                info.OnboardDate = Convert.ToString(sheet.Cells[rowIndex, enterDateColumn].Value); // != null ? sheet.Cells[rowIndex, 10].Value.ToString() : null;

                // leave date
                info.DimissionDate = Convert.ToString(sheet.Cells[rowIndex, leaveDateColumn].Value); // != null ? sheet.Cells[rowIndex, 11].Value.ToString() : null;

                personList.Add(info);
            } // for (rowIndex = 1; rowIndex < 5000; rowIndex++)

            //
            // sort on onboard date by ascend
            personList.Sort(
                (x, y) =>
                    {
                        if (string.IsNullOrWhiteSpace(x.OnboardDate))
                        {
                            return -1; // x - y < 0, put x ahead
                        }
                        return x.OnboardDate.CompareTo(y.OnboardDate);
                    }
                    );

            // 
            // map by name, exclude duplicate records. the latest one will take effect
            personList.ForEach(delegate (PersonInfo info)
                {
                    lock (InfoMap)
                    {
                        InfoMap[info.Name] = info;
                    }
                });

            // correct data and set owner system
            //CorrectData(InfoMap.Values);
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

                if (!string.IsNullOrWhiteSpace(nextPerson.DimissionDate) && (nextPerson.DimissionDate != "长期"))
                {
                    continue;
                }

                outsourceList.Add(nextPerson);
            }

            // sort
            outsourceList.Sort((info1, info2) =>
            {
                if (info1.BizProject != info2.BizProject)
                {
                    return info1.BizProject != null ? info1.BizProject.CompareTo(info2.BizProject) : 1;
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

                if (nextPerson.BizProject != null && nextPerson.BizProject.Contains("大数据"))
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
                if (string.IsNullOrWhiteSpace(nextPerson.BizProject))
                {
                    Trace.WriteLine("nextInfo.Project is empty");
                    continue;
                }

                // correct "消费信贷（一期）"
                if (nextPerson.BizProject.Contains("消费信贷（一期）")
                    || nextPerson.BizProject.Contains("消费信贷（二期）"))
                {
                    nextPerson.BizProject = "消费信贷";
                }

                // 个人信贷系统
                if (nextPerson.BizProject.Contains("消费信贷")
                    || nextPerson.BizProject.Contains("操作平台")
                    || nextPerson.BizProject.Contains("信贷核心")
                    || nextPerson.BizProject.Contains("资信平台")
                    || nextPerson.BizProject.Contains("调度平台"))
                {
                    nextPerson.System = "个人信贷系统";
                    continue;
                }

                // 统一支付平台
                if (nextPerson.BizProject.Contains("统一支付"))
                {
                    nextPerson.System = "统一支付平台";
                    continue;
                }

                // 个人理财系统
                if (nextPerson.BizProject.Contains("个人理财"))
                {
                    nextPerson.System = "个人理财系统";
                    continue;
                }

                // 开放平台
                if (nextPerson.BizProject.Contains("开放平台")
                    || (nextPerson.BizProject.Contains("渠道") && nextPerson.BizProject.Contains("H5")))
                {
                    nextPerson.System = "开放平台";
                    continue;
                }

                // 互金平台
                if (nextPerson.BizProject.Contains("互联网金融平台")
                    || nextPerson.BizProject.Contains("互金平台"))
                {
                    nextPerson.System = "互金平台";
                    continue;
                }

                // 对公信贷系统
                if (nextPerson.BizProject.Contains("对公信贷"))
                {
                    nextPerson.System = "对公信贷系统";
                    continue;
                }

                // 客服系统
                if (nextPerson.BizProject.Contains("客服"))
                {
                    nextPerson.System = "客服系统";
                    continue;
                }

                // 渠道类系统（APP，微信）
                if (nextPerson.BizProject.Contains("APP")
                    || (nextPerson.BizProject.Contains("微信") && (nextPerson.Manager == "王月超"))
                    || (nextPerson.BizProject.Contains("Open") && nextPerson.BizProject.Contains("API"))
                    || (nextPerson.BizProject.Contains("渠道") && nextPerson.BizProject.Contains("API"))
                    || (nextPerson.BizProject.Contains("电子银行渠道") && (nextPerson.Manager == "杨嘉")))
                {
                    nextPerson.System = "渠道类系统";
                    continue;
                }

                // 财管系统
                if (nextPerson.BizProject.Contains("财管")
                    || nextPerson.BizProject.Contains("总账"))
                {
                    nextPerson.System = "财管系统";
                    continue;
                }

                // 促销系统
                if (nextPerson.BizProject.Contains("促销"))
                {
                    nextPerson.System = "促销系统";
                    continue;
                }

                // 大核心银行（含联网核查，电信反诈骗等）
                if (nextPerson.BizProject.Contains("核心银行")
                    || nextPerson.BizProject.Contains("核心系统")
                    || nextPerson.BizProject.Contains("联网核查")
                    || nextPerson.BizProject.Contains("电信反诈骗"))
                {
                    nextPerson.System = "大核心银行";
                    continue;
                }
            }
            #endregion // set owner system
        }
    }
}
