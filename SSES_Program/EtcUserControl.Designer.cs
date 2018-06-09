namespace SSES_Program
{
    partial class EtcUserControl
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
            this.btnFolder = new System.Windows.Forms.Button();
            this.linkLabel_etc = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // btnFolder
            // 
            this.btnFolder.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnFolder.Location = new System.Drawing.Point(311, 1);
            this.btnFolder.Name = "btnFolder";
            this.btnFolder.Size = new System.Drawing.Size(84, 30);
            this.btnFolder.TabIndex = 25;
            this.btnFolder.Text = "폴더보안";
            this.btnFolder.UseVisualStyleBackColor = true;
            this.btnFolder.Visible = false;
            // 
            // linkLabel_etc
            // 
            this.linkLabel_etc.AutoSize = true;
            this.linkLabel_etc.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.linkLabel_etc.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.linkLabel_etc.LinkColor = System.Drawing.Color.Black;
            this.linkLabel_etc.Location = new System.Drawing.Point(16, 9);
            this.linkLabel_etc.Name = "linkLabel_etc";
            this.linkLabel_etc.Size = new System.Drawing.Size(189, 20);
            this.linkLabel_etc.TabIndex = 24;
            this.linkLabel_etc.TabStop = true;
            this.linkLabel_etc.Text = "소비전력 및 전기요금 설정";
            this.linkLabel_etc.Visible = false;
            // 
            // EtcUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Controls.Add(this.btnFolder);
            this.Controls.Add(this.linkLabel_etc);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "EtcUserControl";
            this.Size = new System.Drawing.Size(400, 40);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


     
    }
}
