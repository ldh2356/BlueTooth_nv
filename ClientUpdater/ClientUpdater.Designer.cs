namespace ClientUpdater
{
    partial class ClientUpdater
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
            this.components = new System.ComponentModel.Container();
            this.clientUpdateWizardControl = new DevExpress.XtraWizard.WizardControl();
            this.welcomeWizardPage = new DevExpress.XtraWizard.WelcomeWizardPage();
            this.updateDescrptionWizardPage = new DevExpress.XtraWizard.WizardPage();
            this.updateInformationRichTextBox = new System.Windows.Forms.RichTextBox();
            this.completionWizardPage1 = new DevExpress.XtraWizard.CompletionWizardPage();
            this.downLoadUnpackingWizardPage = new DevExpress.XtraWizard.WizardPage();
            this.updateStatusRichTextBox = new System.Windows.Forms.RichTextBox();
            this.progressBarControl = new DevExpress.XtraEditors.ProgressBarControl();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.clientUpdateWizardControl)).BeginInit();
            this.clientUpdateWizardControl.SuspendLayout();
            this.updateDescrptionWizardPage.SuspendLayout();
            this.downLoadUnpackingWizardPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // clientUpdateWizardControl
            // 
            this.clientUpdateWizardControl.Controls.Add(this.welcomeWizardPage);
            this.clientUpdateWizardControl.Controls.Add(this.updateDescrptionWizardPage);
            this.clientUpdateWizardControl.Controls.Add(this.completionWizardPage1);
            this.clientUpdateWizardControl.Controls.Add(this.downLoadUnpackingWizardPage);
            this.clientUpdateWizardControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clientUpdateWizardControl.Location = new System.Drawing.Point(0, 0);
            this.clientUpdateWizardControl.Name = "clientUpdateWizardControl";
            this.clientUpdateWizardControl.Pages.AddRange(new DevExpress.XtraWizard.BaseWizardPage[] {
            this.welcomeWizardPage,
            this.updateDescrptionWizardPage,
            this.downLoadUnpackingWizardPage,
            this.completionWizardPage1});
            this.clientUpdateWizardControl.ShowHeaderImage = true;
            this.clientUpdateWizardControl.Size = new System.Drawing.Size(598, 323);
            this.clientUpdateWizardControl.Text = "Client Updater";
            this.clientUpdateWizardControl.CancelClick += new System.ComponentModel.CancelEventHandler(this.wizardControl1_CancelClick);
            this.clientUpdateWizardControl.FinishClick += new System.ComponentModel.CancelEventHandler(this.wizardControl1_FinishClick);
            this.clientUpdateWizardControl.NextClick += new DevExpress.XtraWizard.WizardCommandButtonClickEventHandler(this.wizardControl1_NextClick);
            // 
            // welcomeWizardPage
            // 
            this.welcomeWizardPage.IntroductionText = "";
            this.welcomeWizardPage.Name = "welcomeWizardPage";
            this.welcomeWizardPage.Size = new System.Drawing.Size(381, 190);
            this.welcomeWizardPage.Text = "SSES Client Auto Updater";
            this.welcomeWizardPage.Click += new System.EventHandler(this.welcomeWizardPage_Click);
            // 
            // updateDescrptionWizardPage
            // 
            this.updateDescrptionWizardPage.Controls.Add(this.updateInformationRichTextBox);
            this.updateDescrptionWizardPage.DescriptionText = "Show Client Update History";
            this.updateDescrptionWizardPage.Name = "updateDescrptionWizardPage";
            this.updateDescrptionWizardPage.Size = new System.Drawing.Size(566, 178);
            this.updateDescrptionWizardPage.Text = "Update Description";
            // 
            // updateInformationRichTextBox
            // 
            this.updateInformationRichTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.updateInformationRichTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.updateInformationRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.updateInformationRichTextBox.Location = new System.Drawing.Point(0, 0);
            this.updateInformationRichTextBox.Name = "updateInformationRichTextBox";
            this.updateInformationRichTextBox.ReadOnly = true;
            this.updateInformationRichTextBox.Size = new System.Drawing.Size(566, 178);
            this.updateInformationRichTextBox.TabIndex = 2;
            this.updateInformationRichTextBox.Text = "";
            // 
            // completionWizardPage1
            // 
            this.completionWizardPage1.AllowBack = false;
            this.completionWizardPage1.AllowCancel = false;
            this.completionWizardPage1.Name = "completionWizardPage1";
            this.completionWizardPage1.Size = new System.Drawing.Size(381, 190);
            this.completionWizardPage1.Text = "Client Update Finished";
            // 
            // downLoadUnpackingWizardPage
            // 
            this.downLoadUnpackingWizardPage.Controls.Add(this.updateStatusRichTextBox);
            this.downLoadUnpackingWizardPage.Controls.Add(this.progressBarControl);
            this.downLoadUnpackingWizardPage.DescriptionText = "";
            this.downLoadUnpackingWizardPage.Name = "downLoadUnpackingWizardPage";
            this.downLoadUnpackingWizardPage.Size = new System.Drawing.Size(566, 178);
            this.downLoadUnpackingWizardPage.Text = "Client Download And Unpacking";
            // 
            // updateStatusRichTextBox
            // 
            this.updateStatusRichTextBox.BackColor = System.Drawing.SystemColors.Control;
            this.updateStatusRichTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.updateStatusRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.updateStatusRichTextBox.Location = new System.Drawing.Point(0, 0);
            this.updateStatusRichTextBox.Name = "updateStatusRichTextBox";
            this.updateStatusRichTextBox.ReadOnly = true;
            this.updateStatusRichTextBox.Size = new System.Drawing.Size(566, 160);
            this.updateStatusRichTextBox.TabIndex = 5;
            this.updateStatusRichTextBox.Text = "";
            // 
            // progressBarControl
            // 
            this.progressBarControl.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBarControl.EditValue = "20";
            this.progressBarControl.Location = new System.Drawing.Point(0, 160);
            this.progressBarControl.Name = "progressBarControl";
            this.progressBarControl.Properties.ShowTitle = true;
            this.progressBarControl.Size = new System.Drawing.Size(566, 18);
            this.progressBarControl.TabIndex = 4;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // ClientUpdater
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(598, 323);
            this.ControlBox = false;
            this.Controls.Add(this.clientUpdateWizardControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ClientUpdater";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Client Updater";
            this.Load += new System.EventHandler(this.ClientUpdater_Load);
            ((System.ComponentModel.ISupportInitialize)(this.clientUpdateWizardControl)).EndInit();
            this.clientUpdateWizardControl.ResumeLayout(false);
            this.updateDescrptionWizardPage.ResumeLayout(false);
            this.downLoadUnpackingWizardPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraWizard.WizardControl clientUpdateWizardControl;
        private DevExpress.XtraWizard.WelcomeWizardPage welcomeWizardPage;
        private DevExpress.XtraWizard.WizardPage updateDescrptionWizardPage;
        private DevExpress.XtraWizard.CompletionWizardPage completionWizardPage1;
        private DevExpress.XtraWizard.WizardPage downLoadUnpackingWizardPage;
        private System.Windows.Forms.RichTextBox updateInformationRichTextBox;
        private DevExpress.XtraEditors.ProgressBarControl progressBarControl;
        private System.Windows.Forms.RichTextBox updateStatusRichTextBox;
        private System.Windows.Forms.Timer timer1;
    }
}

