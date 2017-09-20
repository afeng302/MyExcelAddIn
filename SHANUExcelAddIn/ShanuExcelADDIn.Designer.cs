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
            this.btnStaffStatistic = new System.Windows.Forms.Button();
            this.btnWorkLoad = new System.Windows.Forms.Button();
            this.btnStatement = new System.Windows.Forms.Button();
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
            this.btnAttendaceExpception.Location = new System.Drawing.Point(3, 89);
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
            this.label1.Size = new System.Drawing.Size(196, 56);
            this.label1.TabIndex = 260;
            this.label1.Text = "目录： “C:\\data”\r\n文件： \r\n    “科技部外包考勤.xlsx”\r\n    “外包人员台账.xlsx”\r\n";
            // 
            // btnStaffStatistic
            // 
            this.btnStaffStatistic.BackColor = System.Drawing.Color.OliveDrab;
            this.btnStaffStatistic.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnStaffStatistic.BackgroundImage")));
            this.btnStaffStatistic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnStaffStatistic.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStaffStatistic.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStaffStatistic.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnStaffStatistic.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnStaffStatistic.Location = new System.Drawing.Point(3, 149);
            this.btnStaffStatistic.Name = "btnStaffStatistic";
            this.btnStaffStatistic.Size = new System.Drawing.Size(117, 43);
            this.btnStaffStatistic.TabIndex = 261;
            this.btnStaffStatistic.Text = "人力外包\r\n人员统计";
            this.btnStaffStatistic.UseVisualStyleBackColor = false;
            this.btnStaffStatistic.Click += new System.EventHandler(this.btnStaffStatistic_Click);
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
            this.btnWorkLoad.Location = new System.Drawing.Point(3, 210);
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
            this.btnStatement.Location = new System.Drawing.Point(3, 270);
            this.btnStatement.Name = "btnStatement";
            this.btnStatement.Size = new System.Drawing.Size(117, 43);
            this.btnStatement.TabIndex = 263;
            this.btnStatement.Text = "结算单";
            this.btnStatement.UseVisualStyleBackColor = false;
            // 
            // ShanuExcelADDIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(172)))), ((int)(((byte)(91)))));
            this.Controls.Add(this.btnStatement);
            this.Controls.Add(this.btnWorkLoad);
            this.Controls.Add(this.btnStaffStatistic);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAttendaceExpception);
            this.Name = "ShanuExcelADDIn";
            this.Size = new System.Drawing.Size(202, 345);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnAttendaceExpception;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnStaffStatistic;
        private System.Windows.Forms.Button btnWorkLoad;
        private System.Windows.Forms.Button btnStatement;
    }
}
