using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EnergySolutions
{
    public delegate void SaveChangedHandler(CalcReduction sender);
    public class CalcReduction
    {
        public SaveChangedHandler OnSaveChanged = null;

        // DevicePerKwh와 WonPerKwh는 MainForm의 생성자에서 AppConfig.xml파일에서 읽어온 값이 들어가게 됨. 저장되있는 값이 없을 경우 기본값으로 160과 183이 적용됨.
        public double DevicePerKwh = 160.0; // 전력값(kW)
        public double WonPerKwh = 183; // won/kWh 전기요금 // MainForm에서 WonPerKwh의 값을 변경하기 위해 double에서 public double로 변경함
        public TimeSpan UsedSec = TimeSpan.FromSeconds(0); // 시간(second)
        public TimeSpan UsedOperation = TimeSpan.FromSeconds(0); // 시간(second)
        double Co2Unit = 0.424;
        double TreeUnit = 2.77;

        public double UsedKwh = 0.0; // 전력량
        public double UsedCost = 0.0; // 사용 전기요금
        public double Co2 = 0.0; // Co2 양
        public double Tree = 0.0; // 나무절약갯수
        public string _IsSend;

        #region filename
        public static string logFileName
        {
            get
            {
                string drivepath = Environment.ExpandEnvironmentVariables("%SystemDrive%") + @"\HansCreative\nv\SSES_Program";
                string fileName = @"\SSES_log.xml";
                return drivepath + fileName;
            }
        }
        #endregion

        public CalcReduction()
        {
            SaveEnergy.Instance.FileName = logFileName;
            SaveEnergy.Instance.LoadFromFile();

            UsedSec = TimeSpan.Parse(SaveEnergy.Instance.UsedSec);
            UsedOperation = TimeSpan.Parse(SaveEnergy.Instance.UsedOperation);

            UsedKwh = SaveEnergy.Instance.UsedKwh; // 전력량
            UsedCost = SaveEnergy.Instance.UsedCost; // 사용 전기요금
            Co2 = SaveEnergy.Instance.Co2; // Co2 양
            Tree = SaveEnergy.Instance.Tree; // 나무절약갯수
            _IsSend = SaveEnergy.Instance.IsSend;      
        }

        public string IsSend
        {
            get
            {
                return _IsSend;
            }
            set
            {
                SaveEnergy.Instance.IsSend = value;
            }
        }

        public CalcReduction(double devicePerKwh)
        {
            DevicePerKwh = devicePerKwh;
        }

       public DateTime _OperationStartTime;

       public DateTime _OperationEndTime;
       public DateTime OperationEndTime
       {
           get
           {
               return _OperationEndTime;
           }

           set
           {
               _OperationEndTime = value;
               TimeSpan t = _OperationEndTime - _OperationStartTime;
               UsedOperation += t;
               SaveEnergy.Instance.UsedOperation = UsedOperation.ToString();
             
               if (OnSaveChanged != null) OnSaveChanged(this);
           }
       }
      
        public DateTime StartTime;

        private DateTime _EndTime;
        public DateTime EndTime
        {
            get
            {
                return _EndTime;
            }

            set
            {
                _EndTime = value;
                TimeSpan t = _EndTime - StartTime;
                UsedSec += t;
                SaveEnergy.Instance.UsedSec = UsedSec.ToString();
                Calculate();
                if (OnSaveChanged != null) OnSaveChanged(this);
            }
        }



        public void Calculate()
        {
            double deviceWatt = DevicePerKwh / 1000; //(kW)

            UsedKwh = deviceWatt * UsedSec.TotalHours;
            UsedCost = UsedKwh * WonPerKwh;
            Co2 = UsedKwh * Co2Unit;
            Tree = Co2 / TreeUnit;

            SaveEnergy.Instance.UsedKwh = UsedKwh;
            SaveEnergy.Instance.UsedCost = UsedCost;
            SaveEnergy.Instance.Co2 = Co2;
            SaveEnergy.Instance.Tree = Tree;
        }
    }
}
