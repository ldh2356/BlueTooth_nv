namespace SSES_Program
{
    partial class DeviceUserControl
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
            this.tbDeviceAddr5 = new System.Windows.Forms.TextBox();
            this.tbDeviceAddr4 = new System.Windows.Forms.TextBox();
            this.tbDeviceAddr3 = new System.Windows.Forms.TextBox();
            this.tbDeviceAddr2 = new System.Windows.Forms.TextBox();
            this.tbDeviceAddr1 = new System.Windows.Forms.TextBox();
            this.btOk = new System.Windows.Forms.Button();
            this.tbDeviceAddr0 = new System.Windows.Forms.TextBox();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tbDeviceAddr5
            // 
            this.tbDeviceAddr5.Location = new System.Drawing.Point(271, 32);
            this.tbDeviceAddr5.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr5.MaxLength = 2;
            this.tbDeviceAddr5.Name = "tbDeviceAddr5";
            this.tbDeviceAddr5.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr5.TabIndex = 15;
            // 
            // tbDeviceAddr4
            // 
            this.tbDeviceAddr4.Location = new System.Drawing.Point(232, 32);
            this.tbDeviceAddr4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr4.MaxLength = 2;
            this.tbDeviceAddr4.Name = "tbDeviceAddr4";
            this.tbDeviceAddr4.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr4.TabIndex = 14;
            // 
            // tbDeviceAddr3
            // 
            this.tbDeviceAddr3.Location = new System.Drawing.Point(191, 32);
            this.tbDeviceAddr3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr3.MaxLength = 2;
            this.tbDeviceAddr3.Name = "tbDeviceAddr3";
            this.tbDeviceAddr3.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr3.TabIndex = 13;
            // 
            // tbDeviceAddr2
            // 
            this.tbDeviceAddr2.Location = new System.Drawing.Point(151, 32);
            this.tbDeviceAddr2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr2.MaxLength = 2;
            this.tbDeviceAddr2.Name = "tbDeviceAddr2";
            this.tbDeviceAddr2.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr2.TabIndex = 12;
            // 
            // tbDeviceAddr1
            // 
            this.tbDeviceAddr1.Location = new System.Drawing.Point(110, 32);
            this.tbDeviceAddr1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr1.MaxLength = 2;
            this.tbDeviceAddr1.Name = "tbDeviceAddr1";
            this.tbDeviceAddr1.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr1.TabIndex = 11;
            // 
            // btOk
            // 
            this.btOk.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.btOk.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btOk.Location = new System.Drawing.Point(309, 29);
            this.btOk.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btOk.Name = "btOk";
            this.btOk.Size = new System.Drawing.Size(84, 27);
            this.btOk.TabIndex = 16;
            this.btOk.Text = "입력";
            this.btOk.UseVisualStyleBackColor = true;
            // 
            // tbDeviceAddr0
            // 
            this.tbDeviceAddr0.Location = new System.Drawing.Point(70, 32);
            this.tbDeviceAddr0.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbDeviceAddr0.MaxLength = 2;
            this.tbDeviceAddr0.Name = "tbDeviceAddr0";
            this.tbDeviceAddr0.Size = new System.Drawing.Size(33, 27);
            this.tbDeviceAddr0.TabIndex = 10;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(14, 8);
            this.radioButton1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(85, 24);
            this.radioButton1.TabIndex = 17;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Android";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(142, 8);
            this.radioButton2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(54, 24);
            this.radioButton2.TabIndex = 18;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "IOS";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 20);
            this.label1.TabIndex = 19;
            this.label1.Text = "장치주소";
            // 
            // DeviceUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.tbDeviceAddr5);
            this.Controls.Add(this.tbDeviceAddr4);
            this.Controls.Add(this.tbDeviceAddr3);
            this.Controls.Add(this.tbDeviceAddr2);
            this.Controls.Add(this.tbDeviceAddr1);
            this.Controls.Add(this.btOk);
            this.Controls.Add(this.tbDeviceAddr0);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "DeviceUserControl";
            this.Size = new System.Drawing.Size(400, 76);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
