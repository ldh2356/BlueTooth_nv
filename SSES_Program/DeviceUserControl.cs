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
    public partial class DeviceUserControl : UserControl
    {
        private System.Windows.Forms.TextBox tbDeviceAddr5;

        public System.Windows.Forms.TextBox TbDeviceAddr5
        {
            get { return tbDeviceAddr5; }
            set { tbDeviceAddr5 = value; }
        }
        private System.Windows.Forms.TextBox tbDeviceAddr4;

        public System.Windows.Forms.TextBox TbDeviceAddr4
        {
            get { return tbDeviceAddr4; }
            set { tbDeviceAddr4 = value; }
        }
        private System.Windows.Forms.TextBox tbDeviceAddr3;

        public System.Windows.Forms.TextBox TbDeviceAddr3
        {
            get { return tbDeviceAddr3; }
            set { tbDeviceAddr3 = value; }
        }
        private System.Windows.Forms.TextBox tbDeviceAddr2;

        public System.Windows.Forms.TextBox TbDeviceAddr2
        {
            get { return tbDeviceAddr2; }
            set { tbDeviceAddr2 = value; }
        }
        private System.Windows.Forms.TextBox tbDeviceAddr1;

        public System.Windows.Forms.TextBox TbDeviceAddr1
        {
            get { return tbDeviceAddr1; }
            set { tbDeviceAddr1 = value; }
        }
        private System.Windows.Forms.Button btOk;

        public System.Windows.Forms.Button BtOk
        {
            get { return btOk; }
            set { btOk = value; }
        }
        private System.Windows.Forms.TextBox tbDeviceAddr0;

        public System.Windows.Forms.TextBox TbDeviceAddr0
        {
            get { return tbDeviceAddr0; }
            set { tbDeviceAddr0 = value; }
        }
        private System.Windows.Forms.RadioButton radioButton1;

        public System.Windows.Forms.RadioButton RadioButton1
        {
            get { return radioButton1; }
            set { radioButton1 = value; }
        }
        private System.Windows.Forms.RadioButton radioButton2;

        public System.Windows.Forms.RadioButton RadioButton2
        {
            get { return radioButton2; }
            set { radioButton2 = value; }
        }
        private System.Windows.Forms.Label label1;

        public System.Windows.Forms.Label Label1
        {
            get { return label1; }
            set { label1 = value; }
        }

        public DeviceUserControl()
        {
            InitializeComponent();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
