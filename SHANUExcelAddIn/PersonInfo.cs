using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHANUExcelAddIn
{
    class PersonInfo
    {
        private string name = string.Empty;
        private string manager = string.Empty;

        public string Name
        {
            get { return this.name != null ? this.name.Trim() : this.name; }
            set { this.name = value; }
        }

        public string Company { get; set; }

        public string Manager
        {
            get { return this.manager != null ? this.manager.Trim() : this.manager; }
            set { this.manager = value; }
        }

        public string Project { get; set; }

        public string Department { get; set; }

        public string System { get; set; }

        public string OnboardDate { get; set; }

        public string DimissionDate { get; set; }

        /// <summary>
        /// 外包模式：项目/人力
        /// </summary>
        public string WorkType { get; set; }
    }
}
