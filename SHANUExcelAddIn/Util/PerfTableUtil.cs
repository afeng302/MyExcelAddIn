using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SHANUExcelAddIn.Util
{
    class KPIInfo
    {
        public string CandidateRole { get; set; }

        public string ID { get; set; }

        public string KPIName { get; set; }

        public string Criteria { get; set; }

        public string EvaluatorRole { get; set; }
    }

    class CandidateInfo
    {
        public string Name { get; set; }

        public string Role { get; set; }

        public string[] ProjectManages { get; set; }

        public string[] TechManagers { get; set; }

        public string[] TestManagers { get; set; }

        public string[] AppOpses { get; set; }

        public string[] DevCenter { get; set; }

        public string[] OpsCenter { get; set; }

        public string[] TestCenter { get; set; }

        public string[] BigDataCenter { get; set; }

        public string[] DevOpsGroup { get; set; }

        public string[] PrdManagers { get; set; }
    }

    class EvaluationInfo
    {
        public string Candidate { get; set; }

        public string KPIID { get; set; }

        public string Evaluator { get; set; }

        public EvaluationInfo Clone()
        {
            EvaluationInfo other = new EvaluationInfo();

            other.Candidate = this.Candidate;
            other.KPIID = this.KPIID;
            other.Evaluator = this.Evaluator;

            return other;
        }
    }

    static class PerfTableUtil
    {
        static List<KPIInfo> KPI_INFO_LIST = new List<KPIInfo>();

        static Dictionary<string, List<KPIInfo>> CANDIDATE_ROLE_KPI_MAP = new Dictionary<string, List<KPIInfo>>();

        static List<CandidateInfo> CANDIDATE_INFO_LIST = new List<CandidateInfo>();

        static List<EvaluationInfo> EVALUATION_INFO_LIST = new List<EvaluationInfo>();

        public static void ReadKPIItems(Excel.Worksheet sheet)
        {

            KPI_INFO_LIST.Clear();
            CANDIDATE_ROLE_KPI_MAP.Clear();
            CANDIDATE_INFO_LIST.Clear();
            EVALUATION_INFO_LIST.Clear();

            #region read out the kpi items
            string currRole = string.Empty;
            for (int rowIndex = 1; rowIndex < 100; rowIndex++)
            {
                KPIInfo Info = new KPIInfo();

                string cellValue = sheet.Cells[rowIndex, 1].Value as string;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    break; // done
                }

                if (cellValue.StartsWith("被考核岗位"))
                {
                    string[] vals = cellValue.Split(new String[] { "：" }, StringSplitOptions.RemoveEmptyEntries);
                    if (vals.Length < 2)
                    {
                        MessageBox.Show("invalid value: " + cellValue);
                        return;
                    }
                    currRole = vals[1];

                    continue;
                }

                if (!Regex.IsMatch(cellValue, "K\\d+", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                Info.CandidateRole = currRole;
                Info.ID = sheet.Cells[rowIndex, 1].Value as string;
                Info.KPIName = sheet.Cells[rowIndex, 2].Value as string;
                Info.Criteria = sheet.Cells[rowIndex, 3].Value as string;
                Info.EvaluatorRole = sheet.Cells[rowIndex, 5].Value as string;

                KPI_INFO_LIST.Add(Info);
            } // for (int rowIndex = 1; rowIndex < 100; rowIndex++)

            #endregion read out the kpi items

            #region group by role

            foreach (var nextItem in KPI_INFO_LIST)
            {
                if (!CANDIDATE_ROLE_KPI_MAP.ContainsKey(nextItem.CandidateRole))
                {
                    CANDIDATE_ROLE_KPI_MAP[nextItem.CandidateRole] = new List<KPIInfo>();
                }
                CANDIDATE_ROLE_KPI_MAP[nextItem.CandidateRole].Add(nextItem);
            }

            #endregion group by role
        }

        public static void ReadCandidates(Excel.Worksheet sheet)
        {
            for (int rowIndex = 2; rowIndex < 200; rowIndex++)
            {
                CandidateInfo info = new CandidateInfo();

                string cellValue = sheet.Cells[rowIndex, 1].Value as string;
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    break;
                }

                info.Name = cellValue;

                cellValue = sheet.Cells[rowIndex, 2].Value as string;
                info.Role = cellValue;

                info.ProjectManages = SplitValues(sheet.Cells[rowIndex, 3].Value as string);

                info.TechManagers = SplitValues(sheet.Cells[rowIndex, 4].Value as string);

                info.TestManagers = SplitValues(sheet.Cells[rowIndex, 5].Value as string);

                info.AppOpses = SplitValues(sheet.Cells[rowIndex, 6].Value as string);

                info.PrdManagers = SplitValues(sheet.Cells[rowIndex, 7].Value as string);

                info.DevCenter = SplitValues(sheet.Cells[rowIndex, 8].Value as string);

                info.OpsCenter = SplitValues(sheet.Cells[rowIndex, 9].Value as string);

                info.TestCenter = SplitValues(sheet.Cells[rowIndex, 10].Value as string);

                info.BigDataCenter = SplitValues(sheet.Cells[rowIndex, 11].Value as string);

                info.DevOpsGroup = SplitValues(sheet.Cells[rowIndex, 12].Value as string);

                CANDIDATE_INFO_LIST.Add(info);
            } // for (int rowIndex = 1; rowIndex < 100; rowIndex++)
        }

        static string[] SplitValues(string val)
        {
            if (string.IsNullOrWhiteSpace(val))
            {
                return null;
            }

            return val.Split(new string[] { "、" }, StringSplitOptions.RemoveEmptyEntries);
        }

        public static void BuildEvaluationInfoList()
        {
            foreach (var nextCandidateInfo in CANDIDATE_INFO_LIST)
            {
                List<EvaluationInfo> evaluationList4OnePerson = new List<EvaluationInfo>();

                List<EvaluationInfo> evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "项目经理", nextCandidateInfo.ProjectManages);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "应用负责人", nextCandidateInfo.TechManagers);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "测试经理", nextCandidateInfo.TestManagers);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "应用运维", nextCandidateInfo.AppOpses);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "产品经理", nextCandidateInfo.PrdManagers);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                    "开发中心", nextCandidateInfo.DevCenter);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                   "运维中心", nextCandidateInfo.OpsCenter);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                   "测试中心", nextCandidateInfo.TestCenter);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                   "大数据中心", nextCandidateInfo.BigDataCenter);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                   "DevOps", nextCandidateInfo.DevOpsGroup);
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                evaluationList = Build4OneEvaluatorRole(nextCandidateInfo.Name, nextCandidateInfo.Role,
                   "【客观量化】", new string[] { "【客观量化】" });
                if (evaluationList != null)
                {
                    evaluationList4OnePerson.AddRange(evaluationList);
                }

                // validate
                ValidateByCandidateRole(nextCandidateInfo.Name, nextCandidateInfo.Role, evaluationList4OnePerson);

                // add to global list
                EVALUATION_INFO_LIST.AddRange(evaluationList4OnePerson);

            } // foreach (var nextCandidateInfo in CANDIDATE_INFO_LIST)
        }

        static List<EvaluationInfo> Build4OneEvaluatorRole(string candidateName, string candidateRole,
            string evaluatorRole, string[] evaluators)
        {
            if (evaluators == null)
            {
                return null;
            }

            // locate candidate role
            if (!CANDIDATE_ROLE_KPI_MAP.ContainsKey(candidateRole))
            {
                MessageBox.Show("cannot find candidat role " + candidateRole + " for " + candidateName);
                return null;
            }

            // locate kpi info
            List<KPIInfo> kpiInfoList = new List<KPIInfo>(); // it is possible for one same role evaluate multiple KPIs
            foreach (var nextInfo in CANDIDATE_ROLE_KPI_MAP[candidateRole])
            {
                if (nextInfo.EvaluatorRole.Contains(evaluatorRole))
                {
                    kpiInfoList.Add(nextInfo);
                }
            }
            if (kpiInfoList.Count == 0)
            {
                MessageBox.Show("cannot find KPI info for evaluator " + evaluatorRole + " for " + candidateName);
                return null;
            }

            // build the evaluator list
            List<EvaluationInfo> infoList = new List<EvaluationInfo>();
            for (int i = 0; i < evaluators.Length; i++)
            {
                foreach (var nextInfo in kpiInfoList)
                {
                    EvaluationInfo info = new EvaluationInfo()
                    {
                        Candidate = candidateName,
                        KPIID = nextInfo.ID
                    };

                    info.Evaluator = evaluators[i];

                    infoList.Add(info);
                }
            }

            return infoList;
        }

        /// <summary>
        /// assure all the KPI items have evaluator
        /// </summary>
        /// <param name="candidateName"></param>
        /// /// <param name="candidateRole"></param>
        /// <param name="evaluationList"></param>
        static void ValidateByCandidateRole(string candidateName, string candidateRole, List<EvaluationInfo> evaluationList)
        {
            // locate candidate role
            if (!CANDIDATE_ROLE_KPI_MAP.ContainsKey(candidateRole))
            {
                MessageBox.Show("cannot find candidat role " + candidateRole);
                return;
            }

            // assure each KPI item has evaluator
            HashSet<string> evaluationKPIHash = new HashSet<string>();
            foreach (var nextInfo in evaluationList)
            {
                evaluationKPIHash.Add(nextInfo.KPIID);
            }
            foreach (var nextKPIInfo in CANDIDATE_ROLE_KPI_MAP[candidateRole])
            {
                if (!evaluationKPIHash.Contains(nextKPIInfo.ID))
                {
                    MessageBox.Show("cannot find evaluator for " + candidateName + " at role " + nextKPIInfo.ID + " / " + nextKPIInfo.EvaluatorRole);
                }
            }
        }

        public static void WriteoutEvaluationList(Excel.Worksheet sheet)
        {
            int rowIndex = 0;
            Microsoft.Office.Interop.Excel.Range objRange = null;

            #region title
            objRange = sheet.Cells[1, 1];
            objRange.Value = "被考核人";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[1, 2];
            objRange.Value = "KPI";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            objRange = sheet.Cells[1, 3];
            objRange.Value = "评价人";
            objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            #endregion

            #region Write out lines
            rowIndex = 2;
            foreach (var nextInfo in EVALUATION_INFO_LIST)
            {
                objRange = sheet.Cells[rowIndex, 1];
                objRange.Value = nextInfo.Candidate;
                objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, 2];
                objRange.Value = nextInfo.KPIID;
                objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                objRange = sheet.Cells[rowIndex, 3];
                objRange.Value = nextInfo.Evaluator;
                objRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                rowIndex++;
            }
            #endregion
        }
    }
}
