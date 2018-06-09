namespace SSES_Program
{
    partial class SettingPopup
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingPopup));
            this.label_Power = new System.Windows.Forms.Label();
            this.comboBox_Pc = new System.Windows.Forms.ComboBox();
            this.textBox_Power = new System.Windows.Forms.TextBox();
            this.label_W = new System.Windows.Forms.Label();
            this.label_Rate = new System.Windows.Forms.Label();
            this.comboBox_Use = new System.Windows.Forms.ComboBox();
            this.textBox_Rate = new System.Windows.Forms.TextBox();
            this.label_WonPerKwh = new System.Windows.Forms.Label();
            this.button_Ok = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label_Power
            // 
            this.label_Power.AutoSize = true;
            this.label_Power.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label_Power.Location = new System.Drawing.Point(12, 17);
            this.label_Power.Name = "label_Power";
            this.label_Power.Size = new System.Drawing.Size(53, 12);
            this.label_Power.TabIndex = 0;
            this.label_Power.Text = "소비전력";
            // 
            // comboBox_Pc
            // 
            this.comboBox_Pc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Pc.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.comboBox_Pc.FormattingEnabled = true;
            this.comboBox_Pc.Location = new System.Drawing.Point(71, 12);
            this.comboBox_Pc.Name = "comboBox_Pc";
            this.comboBox_Pc.Size = new System.Drawing.Size(239, 20);
            this.comboBox_Pc.TabIndex = 1;
            this.comboBox_Pc.SelectedIndexChanged += new System.EventHandler(this.comboBox_Pc_SelectedIndexChanged);
            // 
            // textBox_Power
            // 
            this.textBox_Power.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox_Power.Location = new System.Drawing.Point(316, 11);
            this.textBox_Power.Name = "textBox_Power";
            this.textBox_Power.Size = new System.Drawing.Size(37, 21);
            this.textBox_Power.TabIndex = 2;
            // 
            // label_W
            // 
            this.label_W.AutoSize = true;
            this.label_W.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label_W.Location = new System.Drawing.Point(359, 16);
            this.label_W.Name = "label_W";
            this.label_W.Size = new System.Drawing.Size(52, 12);
            this.label_W.TabIndex = 3;
            this.label_W.Text = "W (Watt)";
            // 
            // label_Rate
            // 
            this.label_Rate.AutoSize = true;
            this.label_Rate.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label_Rate.Location = new System.Drawing.Point(12, 47);
            this.label_Rate.Name = "label_Rate";
            this.label_Rate.Size = new System.Drawing.Size(53, 12);
            this.label_Rate.TabIndex = 5;
            this.label_Rate.Text = "전기요금";
            // 
            // comboBox_Use
            // 
            this.comboBox_Use.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Use.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.comboBox_Use.FormattingEnabled = true;
            this.comboBox_Use.Location = new System.Drawing.Point(71, 44);
            this.comboBox_Use.Name = "comboBox_Use";
            this.comboBox_Use.Size = new System.Drawing.Size(239, 20);
            this.comboBox_Use.TabIndex = 6;
            this.comboBox_Use.SelectedIndexChanged += new System.EventHandler(this.comboBox_Use_SelectedIndexChanged);
            // 
            // textBox_Rate
            // 
            this.textBox_Rate.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.textBox_Rate.Location = new System.Drawing.Point(316, 43);
            this.textBox_Rate.Name = "textBox_Rate";
            this.textBox_Rate.Size = new System.Drawing.Size(37, 21);
            this.textBox_Rate.TabIndex = 7;
            // 
            // label_WonPerKwh
            // 
            this.label_WonPerKwh.AutoSize = true;
            this.label_WonPerKwh.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label_WonPerKwh.Location = new System.Drawing.Point(359, 48);
            this.label_WonPerKwh.Name = "label_WonPerKwh";
            this.label_WonPerKwh.Size = new System.Drawing.Size(46, 12);
            this.label_WonPerKwh.TabIndex = 8;
            this.label_WonPerKwh.Text = "원/kWh";
            // 
            // button_Ok
            // 
            this.button_Ok.Location = new System.Drawing.Point(119, 74);
            this.button_Ok.Name = "button_Ok";
            this.button_Ok.Size = new System.Drawing.Size(75, 23);
            this.button_Ok.TabIndex = 9;
            this.button_Ok.Text = "확인";
            this.button_Ok.UseVisualStyleBackColor = true;
            this.button_Ok.Click += new System.EventHandler(this.button_Ok_Click);
            // 
            // button_Cancel
            // 
            this.button_Cancel.Location = new System.Drawing.Point(200, 74);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(75, 23);
            this.button_Cancel.TabIndex = 10;
            this.button_Cancel.Text = "취소";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // SettingPopup
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(417, 109);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.button_Ok);
            this.Controls.Add(this.label_WonPerKwh);
            this.Controls.Add(this.textBox_Rate);
            this.Controls.Add(this.comboBox_Use);
            this.Controls.Add(this.label_Rate);
            this.Controls.Add(this.label_W);
            this.Controls.Add(this.textBox_Power);
            this.Controls.Add(this.comboBox_Pc);
            this.Controls.Add(this.label_Power);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingPopup";
            this.Text = "소비전력 및 전기요금 설정";
            this.Load += new System.EventHandler(this.SettingPopup_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_Power;
        private System.Windows.Forms.ComboBox comboBox_Pc;
        private System.Windows.Forms.TextBox textBox_Power;
        private System.Windows.Forms.Label label_W;
        private System.Windows.Forms.Label label_Rate;
        private System.Windows.Forms.ComboBox comboBox_Use;
        private System.Windows.Forms.TextBox textBox_Rate;
        private System.Windows.Forms.Label label_WonPerKwh;
        private System.Windows.Forms.Button button_Ok;
        private System.Windows.Forms.Button button_Cancel;
    }
}