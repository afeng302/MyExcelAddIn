namespace SHANUExcelAddIn
{
    partial class ShanuExcelADDIn
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ShanuExcelADDIn));
            this.btnAttendaceExpception = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnWorkHourPoll = new System.Windows.Forms.Button();
            this.btnWorkLoad = new System.Windows.Forms.Button();
            this.btnStatement = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnPerfTable = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAttendaceExpception
            // 
            this.btnAttendaceExpception.BackColor = System.Drawing.Color.OliveDrab;
            this.btnAttendaceExpception.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAttendaceExpception.BackgroundImage")));
            this.btnAttendaceExpception.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAttendaceExpception.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAttendaceExpception.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAttendaceExpception.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnAttendaceExpception.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAttendaceExpception.Location = new System.Drawing.Point(3, 80);
            this.btnAttendaceExpception.Name = "btnAttendaceExpception";
            this.btnAttendaceExpception.Size = new System.Drawing.Size(117, 40);
            this.btnAttendaceExpception.TabIndex = 257;
            this.btnAttendaceExpception.Text = "考勤异常";
            this.btnAttendaceExpception.UseVisualStyleBackColor = false;
            this.btnAttendaceExpception.Click += new System.EventHandler(this.btnAttendanceException_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(3, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(168, 56);
            this.label1.TabIndex = 260;
            this.label1.Text = "目录： C:\\data\r\n文件： \r\n    科技部外包考勤.xlsx\r\n    外包人员台账.xlsx";
            // 
            // btnWorkHourPoll
            // 
            this.btnWorkHourPoll.BackColor = System.Drawing.Color.OliveDrab;
            this.btnWorkHourPoll.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnWorkHourPoll.BackgroundImage")));
            this.btnWorkHourPoll.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnWorkHourPoll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWorkHourPoll.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWorkHourPoll.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnWorkHourPoll.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnWorkHourPoll.Location = new System.Drawing.Point(3, 324);
            this.btnWorkHourPoll.Name = "btnWorkHourPoll";
            this.btnWorkHourPoll.Size = new System.Drawing.Size(117, 43);
            this.btnWorkHourPoll.TabIndex = 261;
            this.btnWorkHourPoll.Text = "工时归集";
            this.btnWorkHourPoll.UseVisualStyleBackColor = false;
            this.btnWorkHourPoll.Click += new System.EventHandler(this.btnWorkHourPoll_Click);
            // 
            // btnWorkLoad
            // 
            this.btnWorkLoad.BackColor = System.Drawing.Color.OliveDrab;
            this.btnWorkLoad.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnWorkLoad.BackgroundImage")));
            this.btnWorkLoad.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnWorkLoad.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWorkLoad.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWorkLoad.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnWorkLoad.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnWorkLoad.Location = new System.Drawing.Point(3, 126);
            this.btnWorkLoad.Name = "btnWorkLoad";
            this.btnWorkLoad.Size = new System.Drawing.Size(117, 43);
            this.btnWorkLoad.TabIndex = 262;
            this.btnWorkLoad.Text = "结算工作量";
            this.btnWorkLoad.UseVisualStyleBackColor = false;
            this.btnWorkLoad.Click += new System.EventHandler(this.btnWorkLoad_Click);
            // 
            // btnStatement
            // 
            this.btnStatement.BackColor = System.Drawing.Color.OliveDrab;
            this.btnStatement.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnStatement.BackgroundImage")));
            this.btnStatement.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnStatement.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStatement.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStatement.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnStatement.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnStatement.Location = new System.Drawing.Point(3, 201);
            this.btnStatement.Name = "btnStatement";
            this.btnStatement.Size = new System.Drawing.Size(117, 43);
            this.btnStatement.TabIndex = 263;
            this.btnStatement.Text = "结算单";
            this.btnStatement.UseVisualStyleBackColor = false;
            this.btnStatement.Click += new System.EventHandler(this.btnStatement_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(3, 182);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(147, 14);
            this.label2.TabIndex = 264;
            this.label2.Text = "文件： 人月单价.xlsx";
            // 
            // btnPerfTable
            // 
            this.btnPerfTable.BackColor = System.Drawing.Color.OliveDrab;
            this.btnPerfTable.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnPerfTable.BackgroundImage")));
            this.btnPerfTable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnPerfTable.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnPerfTable.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPerfTable.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnPerfTable.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPerfTable.Location = new System.Drawing.Point(3, 373);
            this.btnPerfTable.Name = "btnPerfTable";
            this.btnPerfTable.Size = new System.Drawing.Size(117, 43);
            this.btnPerfTable.TabIndex = 265;
            this.btnPerfTable.Text = "绩效考核表";
            this.btnPerfTable.UseVisualStyleBackColor = false;
            this.btnPerfTable.Click += new System.EventHandler(this.btnPerfTable_Click);
            // 
            // ShanuExcelADDIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(172)))), ((int)(((byte)(91)))));
            this.Controls.Add(this.btnPerfTable);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnStatement);
            this.Controls.Add(this.btnWorkLoad);
            this.Controls.Add(this.btnWorkHourPoll);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAttendaceExpception);
            this.Name = "ShanuExcelADDIn";
            this.Size = new System.Drawing.Size(176, 475);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnAttendaceExpception;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnWorkHourPoll;
        private System.Windows.Forms.Button btnWorkLoad;
        private System.Windows.Forms.Button btnStatement;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnPerfTable;
    }
}
