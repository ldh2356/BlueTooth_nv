namespace SSES_Program
{
    partial class PasswordUserControl
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
            this.ChkUserPw = new System.Windows.Forms.Button();
            this.tbUserPw = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // ChkUserPw
            // 
            this.ChkUserPw.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.ChkUserPw.Location = new System.Drawing.Point(311, 8);
            this.ChkUserPw.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ChkUserPw.Name = "ChkUserPw";
            this.ChkUserPw.Size = new System.Drawing.Size(84, 27);
            this.ChkUserPw.TabIndex = 19;
            this.ChkUserPw.Text = "확인";
            this.ChkUserPw.UseVisualStyleBackColor = true;
            // 
            // tbUserPw
            // 
            this.tbUserPw.Location = new System.Drawing.Point(23, 10);
            this.tbUserPw.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbUserPw.Name = "tbUserPw";
            this.tbUserPw.PasswordChar = '*';
            this.tbUserPw.Size = new System.Drawing.Size(166, 27);
            this.tbUserPw.TabIndex = 18;
            // 
            // PasswordUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Controls.Add(this.ChkUserPw);
            this.Controls.Add(this.tbUserPw);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "PasswordUserControl";
            this.Size = new System.Drawing.Size(400, 49);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

    }
}
