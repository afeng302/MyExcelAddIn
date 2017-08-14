using System;
using System.Collections.Generic;
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

        public static void GenerateInfoMap(Excel.Workbook book)
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

                lock (InfoMap)
                {
                    InfoMap[info.Name] = info;
                }

            }
        }
    }
}
