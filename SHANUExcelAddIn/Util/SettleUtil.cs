using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SHANUExcelAddIn.Util
{
    static class SettleUtil
    {
        static Dictionary<string, double> PRICE_MAP = new Dictionary<string, double>();
        static Dictionary<string, double> WORKLOAD_MAP = new Dictionary<string, double>();
        static Dictionary<string, double> PERF_MAP = new Dictionary<string, double>();

        public static void GeneratePriceMap(Excel.Workbook book)
        {
            int sheetIndex = 1;
            Excel.Worksheet sheet = book.Sheets[sheetIndex];
            while ((sheet.Name != "人员单价") && (sheetIndex++ < book.Sheets.Count))
            {
                sheet = book.Sheets[sheetIndex];
            }

            // grades
            int rowIndex = 1;
            string[] grades = new string[20];
            for (int i = 0; i < 20; i++)
            {
                grades[i] = sheet.Cells[rowIndex, i + 1].Value; // 1 base
            }

            // generate map
            for (rowIndex = 1; rowIndex < 100; rowIndex++)
            {
                for (int i = 0; i < 20; i++)
                {
                    string key = GenerateID(sheet.Cells[rowIndex, 1].Value, sheet.Cells[rowIndex, 2].Value, grades[i]);
                    string value = Convert.ToString(sheet.Cells[rowIndex, i + 1].Value);
                    double price = 0;
                    if (double.TryParse(value, out price))
                    {
                        PRICE_MAP[key] = price;
                    }
                }
            } // for (rowIndex = 1; rowIndex < 100; rowIndex++)
        }

        public static double GetUnitPrice(string companyName, string jobType, string grade)
        {
            string key = GenerateID(companyName, jobType, grade);

            if (PRICE_MAP.ContainsKey(key))
            {
                return PRICE_MAP[key];
            }

            return 0;
        }

        public static void GenerateWorkLoadMap(Excel.Workbook book)
        {
            Excel.Worksheet sheet = null;

            if (book.Sheets.Count == 0)
            {
                MessageBox.Show("空文件！");
                return;
            }

            // set the first sheet as default
            sheet = book.Sheets[1];

            foreach (Excel.Worksheet nextSheet in book.Sheets)
            {
                if (nextSheet.Name.Contains("汇总"))
                {
                    sheet = nextSheet;
                    break;
                }
            }

            // identify the column index from first line
            int nameIndex = 0;
            int workloadIndex = 0;
            for (int i = 1; i < 20; i++)
            {
                string value = Convert.ToString(sheet.Cells[1, i].Value);
                if (string.IsNullOrWhiteSpace(value))
                {
                    break;
                }

                if (value.Contains("姓名"))
                {
                    nameIndex = i;
                }
                else if (value.Contains("结算") && value.Contains("人月"))
                {
                    workloadIndex = i;
                }
            }

            if (nameIndex == 0)
            {
                MessageBox.Show("没有找到【姓名】");
                return;
            }

            if (workloadIndex == 0)
            {
                MessageBox.Show("没有找到【结算人月】");
                return;
            }

            for (int rowIndex = 2; rowIndex < 1000; rowIndex++)
            {
                // 姓名
                string name = sheet.Cells[rowIndex, nameIndex].Value;
                if (string.IsNullOrWhiteSpace(name))
                {
                    break;
                }

                // 结算人月
                string value = Convert.ToString(sheet.Cells[rowIndex, workloadIndex].Value);
                double workload = 0;
                double.TryParse(value, out workload);

                WORKLOAD_MAP[name] = workload;
            }
        }

        public static void GeneratePerformanceMap(Excel.Workbook book)
        {
            Excel.Worksheet sheet = null;

            if (book.Sheets.Count == 0)
            {
                MessageBox.Show("空文件");
                return;
            }

            // set the first sheet as default
            sheet = book.Sheets[1];

            foreach (Excel.Worksheet nextSheet in book.Sheets)
            {
                if (nextSheet.Name.Contains("汇总"))
                {
                    sheet = nextSheet;
                    break;
                }
            }

            // locate the "姓名" row
            int nameRowIndex = 1;
            string value = Convert.ToString(sheet.Cells[nameRowIndex, 1].Value);
            while (!string.IsNullOrWhiteSpace(value) && !value.Contains("姓名"))
            {
                nameRowIndex++;
                value = Convert.ToString(sheet.Cells[nameRowIndex, 1].Value);
            }
            if (string.IsNullOrWhiteSpace(value))
            {
                MessageBox.Show("没有找到【姓名】");
                return;
            }

            // locate the "绩效系数" column
            int perfColIndex = 0;
            for (int i = 1; i < 20; i++)
            {
                value = Convert.ToString(sheet.Cells[nameRowIndex, i].Value);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (value.Contains("绩效") && value.Contains("系数"))
                {
                    perfColIndex = i;
                    break;
                }
            }
            if (perfColIndex == 0)
            {
                MessageBox.Show("没有找到【绩效系数】");
                return; 
            }

            // move to next row
            nameRowIndex++;
            for (int i = nameRowIndex; i < 5000; i++)
            {
                // 姓名
                string name = Convert.ToString(sheet.Cells[i, 1].Value);
                if (string.IsNullOrWhiteSpace(name))
                {
                    if (i > 20)
                    {
                        break; // has complete all the rows in this sheet
                    }
                    continue; // maybe have more rows later
                }

                // 绩效
                double perf = 1.0;
                try
                {
                    perf = Convert.ToDouble(sheet.Cells[i, perfColIndex].Value);
                }
                catch (Exception)
                {
                    value = Convert.ToString(sheet.Cells[i, perfColIndex].Value);
                    System.Diagnostics.Trace.WriteLine("invlid performance value: " + value);
                }


                PERF_MAP[name] = perf;
            } // while (!string.IsNullOrWhiteSpace(value))
        }

        public static double GetPerformance(string name)
        {
            if (PERF_MAP.ContainsKey(name))
            {
                return PERF_MAP[name];
            }
            return 0;
        }

        public static List<string> GetNameList()
        {
            return WORKLOAD_MAP.Keys.ToList();
        }

        public static double GetSettlementMonth(string name)
        {
            if (WORKLOAD_MAP.ContainsKey(name))
            {
                return WORKLOAD_MAP[name];
            }

            return 0;
        }

        static string GenerateID(string companyName, string jobType, string grade)
        {
            return string.Format("[{0}]-[{1}]-[{2}]", companyName, jobType, grade);
        }
    }
}
