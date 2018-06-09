using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SSES_Program
{
    public partial class SettingPopup : Form
    {
        private MainForm main_frm;
        private double power; // 소비전력 (kW)
        private double rate; // 전기요금 (원/kWh)

        public delegate void Popup_EventHandler(double power, double rate);
        public event Popup_EventHandler Popup_BtnClickEvent;

        public SettingPopup()
        {
            InitializeComponent();
        }

        private void SettingPopup_Load(object sender, EventArgs e)
        {
            main_frm = (MainForm)this.Owner;

            textBox_Power.Text = AppConfig.Instance.PcPower.ToString();
            textBox_Rate.Text = AppConfig.Instance.ElecRate.ToString();

            if (main_frm.checkBox_LangToggle.Text.Equals("English"))
            {
                //
                this.Text = "소비전력 및 전기요금 설정";
                label_Power.Text = "소비전력";
                label_Rate.Text = "전기요금";
                label_WonPerKwh.Text = "원/kWh";
                button_Ok.Text = "확인";
                button_Cancel.Text = "취소";

                // comboBox_Pc
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X3K-K57D)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5L-KSF)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5L-K99)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5M-K58)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z54L)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z54M)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5H-K50D)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z77L)");
                comboBox_Pc.Items.Add("LG gram");
                comboBox_Pc.Items.Add("DELL");
                comboBox_Pc.Items.Add("사용자설정");

                // comboBox_Use
                comboBox_Use.Items.Add("주택용(저압)");
                comboBox_Use.Items.Add("주택용(고압)");
                comboBox_Use.Items.Add("일반용(갑) 1");
                comboBox_Use.Items.Add("일반용(갑) 2");
                comboBox_Use.Items.Add("일반용(을)");
                comboBox_Use.Items.Add("산업용(갑) 1");
                comboBox_Use.Items.Add("산업용(갑) 2");
                comboBox_Use.Items.Add("산업용(을)");
                comboBox_Use.Items.Add("교육용(갑)");
                comboBox_Use.Items.Add("교육용(을)");
                comboBox_Use.Items.Add("사용자설정(원)");
                //comboBox_Use.Items.Add("사용자설정(USD)");

            }
            else
            {
                //
                this.Text = "Set power consumption and electricity rate";
                label_Power.Text = "Power";
                label_Rate.Text = "Rate";
                label_WonPerKwh.Text = "USD/kWh";
                button_Ok.Text = "OK";
                button_Cancel.Text = "Cancel";

                // comboBox_Pc
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X3K-K57D)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5L-KSF)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5L-K99)");
                comboBox_Pc.Items.Add("Samsung notebook9 (NT900X5M-K58)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z54L)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z54M)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5H-K50D)");
                comboBox_Pc.Items.Add("Samsung notebook5 (NT500R5L-Z77L)");
                comboBox_Pc.Items.Add("LG gram");
                comboBox_Pc.Items.Add("DELL");
                comboBox_Pc.Items.Add("User Setting");

                // comboBox_Use
                //comboBox_Use.Items.Add("User Setting(KRW)");
                comboBox_Use.Items.Add("User Setting(USD)");
            }

            // 기존 설정을 가져옴
            textBox_Power.Text = AppConfig.Instance.PcPower.ToString();
            textBox_Rate.Text = AppConfig.Instance.ElecRate.ToString();

            textBox_Power.Enabled = false;
            textBox_Rate.Enabled = false;
        }

        private void comboBox_Pc_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox_Power.Enabled = false;

            switch (comboBox_Pc.SelectedIndex)
            {
                case 0:
                    textBox_Power.Text = "130";
                    break;
                case 1:
                    textBox_Power.Text = "293";
                    break;
                case 2:
                    textBox_Power.Text = "293";
                    break;
                case 3:
                    textBox_Power.Text = "170";
                    break;
                case 4:
                    textBox_Power.Text = "246";
                    break;
                case 5:
                    textBox_Power.Text = "246";
                    break;
                case 6:
                    textBox_Power.Text = "256";
                    break;
                case 7:
                    textBox_Power.Text = "246";
                    break;
                case 8:
                    textBox_Power.Text = "40";
                    break;
                case 9:
                    textBox_Power.Text = "45";
                    break;
                case 10:
                    textBox_Power.Enabled = true;
                    textBox_Power.Clear();
                    textBox_Power.Focus();
                    break;
                default:
                    break;
            }
        }

        private void comboBox_Use_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox_Rate.Enabled = false;

            if (main_frm.checkBox_LangToggle.Text.Equals("English"))
            {
                label_WonPerKwh.Text = "원/kWh";

                switch (comboBox_Use.SelectedIndex)
                {
                    case 0:
                        textBox_Rate.Text = "187.9"; // 주택용(저압)
                        break;
                    case 1:
                        textBox_Rate.Text = "147.3"; // 주택용(고압)
                        break;
                    case 2:
                        textBox_Rate.Text = "92.46"; // 일반용(갑) 1
                        break;
                    case 3:
                        textBox_Rate.Text = "86.38"; // 일반용(갑) 2
                        break;
                    case 4:
                        textBox_Rate.Text = "104.43"; // 일반용(을)
                        break;
                    case 5:
                        textBox_Rate.Text = "77.37"; // 산업용(갑) 1
                        break;
                    case 6:
                        textBox_Rate.Text = "79.14"; // 산업용(갑) 2
                        break;
                    case 7:
                        textBox_Rate.Text = "104.43"; // 산업용(을)
                        break;
                    case 8:
                        textBox_Rate.Text = "77.73"; // 교육용(갑)
                        break;
                    case 9:
                        textBox_Rate.Text = "83.27"; // 교육용(을)
                        break;
                    case 10:
                        textBox_Rate.Enabled = true; // 사용자설정(원)
                        textBox_Rate.Clear();
                        textBox_Rate.Focus();
                        break;
                    default:
                        break;
                }
            }
            else
            {
                switch (comboBox_Use.SelectedIndex)
                {
                    case 0:
                        textBox_Rate.Enabled = true; // 사용자설정(USD)
                        label_WonPerKwh.Text = "USD/kWh";
                        textBox_Rate.Clear();
                        textBox_Rate.Focus();
                        break;
                    default:
                        break;
                }
            }
            
        }

        private void button_Ok_Click(object sender, EventArgs e)
        {
            if (textBox_Power.Text != "")
            {
                if (textBox_Rate.Text != "")
                {
                    power = Convert.ToDouble(textBox_Power.Text);
                    rate = Convert.ToDouble(textBox_Rate.Text);

                    Popup_BtnClickEvent(power, rate);
                    this.Close();
                }
                else
                {
                    if (main_frm.checkBox_LangToggle.Text.Equals("English"))
                        MessageBox.Show("전기요금을 입력해주세요", "전기요금 재입력");
                    else
                        MessageBox.Show("Please input your electricity rate", "Input Error");
                }
            }
            else
            {
                if (main_frm.checkBox_LangToggle.Text.Equals("English"))
                    MessageBox.Show("소비전력을 입력해주세요", "소비전력 재입력");
                else
                    MessageBox.Show("Please input your power consumption", "Input Error");
            }
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
