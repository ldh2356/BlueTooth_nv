namespace SSES_Program
{
    partial class SleepModeUserControl
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
            this.ChkMode = new System.Windows.Forms.Button();
            this.rbPcMode = new System.Windows.Forms.RadioButton();
            this.rbMonitorMode = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // ChkMode
            // 
            this.ChkMode.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.ChkMode.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ChkMode.Location = new System.Drawing.Point(311, 9);
            this.ChkMode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ChkMode.Name = "ChkMode";
            this.ChkMode.Size = new System.Drawing.Size(84, 47);
            this.ChkMode.TabIndex = 16;
            this.ChkMode.Text = "확인";
            this.ChkMode.UseVisualStyleBackColor = true;
            // 
            // rbPcMode
            // 
            this.rbPcMode.AutoSize = true;
            this.rbPcMode.BackColor = System.Drawing.Color.Transparent;
            this.rbPcMode.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.rbPcMode.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.rbPcMode.Location = new System.Drawing.Point(33, 33);
            this.rbPcMode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rbPcMode.Name = "rbPcMode";
            this.rbPcMode.Size = new System.Drawing.Size(151, 24);
            this.rbPcMode.TabIndex = 15;
            this.rbPcMode.TabStop = true;
            this.rbPcMode.Text = "모니터+본체 절전";
            this.rbPcMode.UseVisualStyleBackColor = true;
            // 
            // rbMonitorMode
            // 
            this.rbMonitorMode.AutoSize = true;
            this.rbMonitorMode.BackColor = System.Drawing.Color.Transparent;
            this.rbMonitorMode.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.rbMonitorMode.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.rbMonitorMode.Location = new System.Drawing.Point(33, 10);
            this.rbMonitorMode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rbMonitorMode.Name = "rbMonitorMode";
            this.rbMonitorMode.Size = new System.Drawing.Size(110, 24);
            this.rbMonitorMode.TabIndex = 14;
            this.rbMonitorMode.TabStop = true;
            this.rbMonitorMode.Text = "모니터 절전";
            this.rbMonitorMode.UseVisualStyleBackColor = false;
            // 
            // SleepModeUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Controls.Add(this.ChkMode);
            this.Controls.Add(this.rbPcMode);
            this.Controls.Add(this.rbMonitorMode);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SleepModeUserControl";
            this.Size = new System.Drawing.Size(400, 66);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

   
    }
}
