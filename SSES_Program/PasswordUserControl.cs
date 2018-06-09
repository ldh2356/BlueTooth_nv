using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SSES_Program
{
    public partial class PasswordUserControl : UserControl
    {

        private System.Windows.Forms.Button ChkUserPw;

        public System.Windows.Forms.Button ChkUserPw1
        {
            get { return ChkUserPw; }
            set { ChkUserPw = value; }
        }
        private System.Windows.Forms.TextBox tbUserPw;

        public System.Windows.Forms.TextBox TbUserPw
        {
            get { return tbUserPw; }
            set { tbUserPw = value; }
        }

        public PasswordUserControl()
        {
            InitializeComponent();
        }
    }
}
