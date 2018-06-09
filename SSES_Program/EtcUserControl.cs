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
    public partial class EtcUserControl : UserControl
    {
        private System.Windows.Forms.Button btnFolder;
        private System.Windows.Forms.LinkLabel linkLabel_etc;

        public System.Windows.Forms.Button BtnFolder
        {
            get { return btnFolder; }
            set { btnFolder = value; }
        }
        public System.Windows.Forms.LinkLabel LinkLabel_etc
        {
            get { return linkLabel_etc; }
            set { linkLabel_etc = value; }
        }

        public EtcUserControl()
        {
            InitializeComponent();
        }
    }
}
