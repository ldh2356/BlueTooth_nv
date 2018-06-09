using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace SSES_Program
{
    public partial class FormScreenSaverCancel : Form
    {
        FormScreenSaver formScreenSaver;
        FormScreenSaver2 formScreenSaver2;

        public FormScreenSaverCancel()
        {
            InitializeComponent();
        }

        public FormScreenSaverCancel(FormScreenSaver formScreenSaver)
        {
            InitializeComponent();
            // 왜 있어야 하는지 모르겠음
            
            this.formScreenSaver = formScreenSaver;
            formScreenSaver.SetFormScreenSaverCancel(this); // 없어도 되나? 가보면 주석처리 돼있음

            this.TopMost = true;
            
        }

        
        // 새로 만든거
        // (FormScreenSaver2 에서 클릭 이벤트를 받았을 때 비밀번호 입력 창을 띄워주기 위한 생성자)
        public FormScreenSaverCancel(FormScreenSaver2 formScreenSaver2)
        {
            InitializeComponent();

            this.formScreenSaver2 = formScreenSaver2;
            //formScreenSaver2.SetFormScreenSaverCancel(this);

            this.TopMost = true;
        }
        

        // 취소 버튼
        /*
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        */

        // 확인 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            ButtonEvent();
        }

        // 비번 박스
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ButtonEvent();
            }
        }

        void ButtonEvent()
        {
            //if (textBox1.Text == "0000")
            if (textBox1.Text == MainForm.userPw)
            {
                this.Close();
                //formScreenSaver.main.screenSaverAllStop(); // 기존에 이거만 있었음

                // 여기부터 내가 추가한거
                // 1. 스크린 세이버만 종료
                Screen[] screen = Screen.AllScreens; // 시스템 내 모든 디스플레이 배열을 가져옴
                
                if (screen.GetLength(0) != 2) // 듀얼 모니터가 아닌 경우
                    formScreenSaver.main.screenSaverAllStop();
                else // 듀얼 모니터인 경우
                    formScreenSaver2.main.screenSaverAllStop();
            }
            else
            {
                Service.AlertSoundStart();
                MessageBox.Show("비밀번호가 틀렸습니다. \n다시 입력해 주세요.");
            }
        }
    }
}
