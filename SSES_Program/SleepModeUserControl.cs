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
    public partial class SleepModeUserControl : UserControl
    {
        private System.Windows.Forms.Button ChkMode;

        public System.Windows.Forms.Button ChkMode1
        {
            get { return ChkMode; }
            set { ChkMode = value; }
        }
        private System.Windows.Forms.RadioButton rbPcMode;

        public System.Windows.Forms.RadioButton RbPcMode
        {
            get { return rbPcMode; }
            set { rbPcMode = value; }
        }
        private System.Windows.Forms.RadioButton rbMonitorMode;

        public System.Windows.Forms.RadioButton RbMonitorMode
        {
            get { return rbMonitorMode; }
            set { rbMonitorMode = value; }
        }

        public SleepModeUserControl()
        {
            InitializeComponent();
        }
    }
}
