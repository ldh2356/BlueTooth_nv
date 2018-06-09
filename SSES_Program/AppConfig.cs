using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Security.Cryptography;

namespace SSES_Program
{
    public class AppConfig
    {
        private static AppConfig instance = new AppConfig();
        public static AppConfig Instance
        {
            get { return instance; }
        }

        public string FileName;

        public string _LocalName;
        public int _TrackBar;
        public double _PcPower;
        public double _ElecRate; // 추가함
        public int _TotalTime;
        public string _UserPassword;
        public string _DeviceAddress;
        public int _SleepMode;
        public bool _MonitorSleepMode;
        public bool _PcSleepMode;
        private int _Model;

        public int Model
        {
            get { return _Model; }
            set { _Model = value; SaveToFile(); }
        }

        public string LocalName
        {
            get { return _LocalName; }
            set { _LocalName = value; SaveToFile(); }
        }
        public int TrackBar
        {
            get { return _TrackBar; }
            set { _TrackBar = value; SaveToFile(); }
        }

        public double PcPower
        {
            get { return _PcPower; }
            set { _PcPower = value; SaveToFile(); }
        }
        
        public double ElecRate // 추가함
        {
            get { return _ElecRate; }
            set { _ElecRate = value; SaveToFile(); }
        }

        public int TotalTime
        {
            get { return _TotalTime; }
            set { _TotalTime = value; SaveToFile(); }
        }

        public string UserPassword
        {
            get { return _UserPassword; }
            set { _UserPassword = value; SaveToFile(); }
        }

        public string DeviceAddress
        {
            get { return _DeviceAddress; }
            set { _DeviceAddress = value; SaveToFile(); }
        }

        public int SleepMode
        {
            get { return _SleepMode; }
            set { _SleepMode = value; SaveToFile(); }
        }

        public bool MonitorSleepMode
        {
            get { return _MonitorSleepMode; }
            set { _MonitorSleepMode = value; SaveToFile(); }
        }

        public bool PcSleepMode
        {
            get { return _PcSleepMode; }
            set { _PcSleepMode = value; SaveToFile(); }
        }

        public AppConfig()
        {
            init();
        }

        void init()
        {
            _LocalName = "";
            _TrackBar = -65;
            _PcPower = 160; // PC 소비전력 기본값
            _ElecRate = 183; // 전기요금 기본값
            _TotalTime = 0;
            _UserPassword = "0000";
            _DeviceAddress = "00:00:00:00:00:00";
            _SleepMode = 0;
            _Model = 0;
            //_MonitorSleepMode = true;
            //_PcSleepMode = false;
        }

        //암호화 키.  8글자로 이루어짐.
        static byte[] Skey = ASCIIEncoding.ASCII.GetBytes("11111111");

        public void SaveToFile(string fileName = null)
        {
            if (fileName == null) fileName = FileName;

            XElement xe = new XElement("AppConfig");

            xe.Add(new XElement("LocalName", LocalName));
            xe.Add(new XElement("TrackBar", TrackBar));
            xe.Add(new XElement("PcPower", PcPower));
            xe.Add(new XElement("TotalTime", TotalTime));
            xe.Add(new XElement("UserPassword", UserPassword));
            xe.Add(new XElement("DeviceAddress", DeviceAddress));
            xe.Add(new XElement("SleepMode", SleepMode));
            //xe.Add(new XElement("MonitorSleepMode", MonitorSleepMode));
            //xe.Add(new XElement("PcSleepMode", PcSleepMode));
            xe.Add(new XElement("ElecRate", ElecRate)); // 추가함
            xe.Add(new XElement("Model", Model));


            xe.Save(fileName);
            Console.WriteLine("SSES_config.xml ==>" + xe.ToString());

            byte[] data = File.ReadAllBytes(fileName);
            //encrypt
            DESCryptoServiceProvider rc2 = new DESCryptoServiceProvider();
            rc2.Key = Skey;
            rc2.IV = Skey;

            MemoryStream ms = new MemoryStream();

            CryptoStream cryStream = new CryptoStream(ms, rc2.CreateEncryptor(), CryptoStreamMode.Write);
            cryStream.Write(data, 0, data.Length);
            cryStream.FlushFinalBlock();

            byte[] data1 = ms.ToArray();

            File.WriteAllBytes(fileName, data1);

        }

        public void LoadFromFile(string fileName = null)
        {
            if (fileName == null) fileName = FileName;
            if (!File.Exists(fileName)) return;

            byte[] data = File.ReadAllBytes(fileName);
            //decrypt
            DESCryptoServiceProvider rc2 = new DESCryptoServiceProvider();
            rc2.Key = Skey;
            rc2.IV = Skey;
            MemoryStream ms = new MemoryStream();

            CryptoStream cryStream = new CryptoStream(ms, rc2.CreateDecryptor(), CryptoStreamMode.Write);

            cryStream.Write(data, 0, data.Length);
            try
            {
                cryStream.FlushFinalBlock();
            }
            catch (Exception ex)
            {
                Console.WriteLine("D + {0}", ex.Message);
            }

            byte[] data1 = ms.ToArray();
            File.WriteAllBytes(fileName, data1);

            XElement xe = XElement.Load(fileName);
            
            _LocalName = xe.Element("LocalName").Value;
            _TrackBar = int.Parse(xe.Element("TrackBar").Value);
            _PcPower = double.Parse(xe.Element("PcPower").Value);
            _TotalTime = int.Parse(xe.Element("TotalTime").Value);
            _UserPassword = xe.Element("UserPassword").Value;
            _DeviceAddress = xe.Element("DeviceAddress").Value;
            _SleepMode = int.Parse(xe.Element("SleepMode").Value);
            _Model = int.Parse(xe.Element("Model").Value);

            // 추가함
            try
            {
                _ElecRate = double.Parse(xe.Element("ElecRate").Value);
            }
            catch (Exception ex)
            {
                ex.ToString();
                xe.Add(new XElement("ElecRate", ElecRate));
                _ElecRate = double.Parse(xe.Element("ElecRate").Value);
            }
        }
    }

}
