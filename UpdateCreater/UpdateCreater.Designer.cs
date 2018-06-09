namespace UpdateCreater
{
    partial class UpdateCreater
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

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.btnUpload = new DevExpress.XtraEditors.SimpleButton();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.newVersionRichTextBox = new System.Windows.Forms.RichTextBox();
            this.historyRichTextBox = new System.Windows.Forms.RichTextBox();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // textEdit1
            // 
            this.textEdit1.Location = new System.Drawing.Point(41, 15);
            this.textEdit1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textEdit1.Name = "textEdit1";
            this.textEdit1.Properties.ReadOnly = true;
            this.textEdit1.Size = new System.Drawing.Size(251, 24);
            this.textEdit1.TabIndex = 0;
            // 
            // simpleButton1
            // 
            this.simpleButton1.Location = new System.Drawing.Point(299, 12);
            this.simpleButton1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.Size = new System.Drawing.Size(86, 29);
            this.simpleButton1.TabIndex = 1;
            this.simpleButton1.Text = "파일 찾기";
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(7, 19);
            this.labelControl1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(26, 18);
            this.labelControl1.TabIndex = 2;
            this.labelControl1.Text = "경로";
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(14, 369);
            this.btnUpload.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(371, 46);
            this.btnUpload.TabIndex = 3;
            this.btnUpload.Text = "파일 업로드";
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(14, 59);
            this.labelControl2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(70, 18);
            this.labelControl2.TabIndex = 4;
            this.labelControl2.Text = "번경된 사항";
            // 
            // newVersionRichTextBox
            // 
            this.newVersionRichTextBox.Location = new System.Drawing.Point(14, 84);
            this.newVersionRichTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.newVersionRichTextBox.Name = "newVersionRichTextBox";
            this.newVersionRichTextBox.Size = new System.Drawing.Size(371, 122);
            this.newVersionRichTextBox.TabIndex = 5;
            this.newVersionRichTextBox.Text = "";
            // 
            // historyRichTextBox
            // 
            this.historyRichTextBox.Location = new System.Drawing.Point(14, 239);
            this.historyRichTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.historyRichTextBox.Name = "historyRichTextBox";
            this.historyRichTextBox.ReadOnly = true;
            this.historyRichTextBox.Size = new System.Drawing.Size(371, 122);
            this.historyRichTextBox.TabIndex = 6;
            this.historyRichTextBox.Text = "";
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(14, 214);
            this.labelControl3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(52, 18);
            this.labelControl3.TabIndex = 7;
            this.labelControl3.Text = "히스토리";
            // 
            // UpdateCreater
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 434);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.historyRichTextBox);
            this.Controls.Add(this.newVersionRichTextBox);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.simpleButton1);
            this.Controls.Add(this.textEdit1);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "UpdateCreater";
            this.Text = "UpdateCreater";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.UpdateCreater_FormClosing);
            this.Load += new System.EventHandler(this.UpdateCreater_Load);
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.TextEdit textEdit1;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.SimpleButton btnUpload;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private System.Windows.Forms.RichTextBox newVersionRichTextBox;
        private System.Windows.Forms.RichTextBox historyRichTextBox;
        private DevExpress.XtraEditors.LabelControl labelControl3;
    }
}

