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
            this.btnAddText = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnAddText
            // 
            this.btnAddText.BackColor = System.Drawing.Color.OliveDrab;
            this.btnAddText.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAddText.BackgroundImage")));
            this.btnAddText.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAddText.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAddText.Font = new System.Drawing.Font("Tahoma", 16F, System.Drawing.FontStyle.Bold);
            this.btnAddText.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.btnAddText.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddText.Location = new System.Drawing.Point(3, 89);
            this.btnAddText.Name = "btnAddText";
            this.btnAddText.Size = new System.Drawing.Size(117, 40);
            this.btnAddText.TabIndex = 257;
            this.btnAddText.Text = "考勤异常";
            this.btnAddText.UseVisualStyleBackColor = false;
            this.btnAddText.Click += new System.EventHandler(this.btnAddText_Click);
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
            // ShanuExcelADDIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(172)))), ((int)(((byte)(91)))));
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAddText);
            this.Name = "ShanuExcelADDIn";
            this.Size = new System.Drawing.Size(202, 249);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnAddText;
        private System.Windows.Forms.Label label1;
    }
}
