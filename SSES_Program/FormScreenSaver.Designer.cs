namespace SSES_Program
{
    partial class FormScreenSaver
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
            this.pb_screenSaver = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pb_screenSaver)).BeginInit();
            this.SuspendLayout();
            // 
            // pb_screenSaver
            // 
            this.pb_screenSaver.Image = global::SSES_Program.Properties.Resources.ScreenSaver_2017_03_24;
            this.pb_screenSaver.Location = new System.Drawing.Point(0, 0);
            this.pb_screenSaver.Name = "pb_screenSaver";
            this.pb_screenSaver.Size = new System.Drawing.Size(286, 261);
            this.pb_screenSaver.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pb_screenSaver.TabIndex = 0;
            this.pb_screenSaver.TabStop = false;
            this.pb_screenSaver.Click += new System.EventHandler(this.pb_screenSaver_Click);
            this.pb_screenSaver.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pb_screenSaver_MouseDown);
            // 
            // FormScreenSaver
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.ControlBox = false;
            this.Controls.Add(this.pb_screenSaver);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormScreenSaver";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "ScreenSaver";
            this.TopMost = true;
            this.Activated += new System.EventHandler(this.FormScreenSaver_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormScreenSaver_FormClosing);
            this.Load += new System.EventHandler(this.FormScreenSaver_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormScreenSaver_KeyDown);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pb_screenSaver_MouseDown);
            ((System.ComponentModel.ISupportInitialize)(this.pb_screenSaver)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.PictureBox pb_screenSaver;
    }
}