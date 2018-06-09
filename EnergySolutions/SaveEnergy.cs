using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Security.Cryptography;

namespace EnergySolutions
{
    class SaveEnergy
    {
        private static SaveEnergy instance = new SaveEnergy();
        public static SaveEnergy Instance
        {
            get { return instance; }
        }

        public string FileName;

        public TimeSpan _UsedSec; // 시간(second)
        public TimeSpan _UsedOperation;
        public double _UsedKwh; // 전력량
        public double _UsedCost; // 사용 전기요금
        public double _Co2; // Co2 양
        public double _Tree; // 나무절약갯수
        public string _IsSend;

        public string UsedSec
        {
            get { return _UsedSec.ToString(); }
            set { _UsedSec = TimeSpan.Parse(value); SaveToFile(); }
        }

        public string UsedOperation
        {
            get { return _UsedOperation.ToString(); }
            set { _UsedOperation = TimeSpan.Parse(value); SaveToFile(); }
        }

        public double UsedKwh
        {
            get { return _UsedKwh; }
            set { _UsedKwh = value; SaveToFile(); }
        }
        public double UsedCost
        {
            get { return _UsedCost; }
            set { _UsedCost = value; SaveToFile(); }
        }
        public double Co2
        {
            get { return _Co2; }
            set { _Co2 = value; SaveToFile(); }
        }
        public double Tree
        {
            get { return _Tree; }
            set { _Tree = value; SaveToFile(); }
        }

        public string IsSend
        {
            get { return _IsSend; }
            set { _IsSend = value; SaveToFile(); }
        }

        public SaveEnergy()
        {
            init();
        }

        void init()
        {
            _UsedSec = TimeSpan.FromSeconds(0) ; // 시간(second)
            _UsedKwh = 0.0; // 전력량
            _UsedCost = 0.0; // 사용 전기요금
            _Co2 = 0.0; // Co2 양
            _Tree = 0.0; // 나무절약갯수
        }

        public void SaveToFile(string fileName = null)
        {
            if (fileName == null) fileName = FileName;

            XElement xe = new XElement("SaveEnergy");

            xe.Add(new XElement("UsedSec", UsedSec));
            xe.Add(new XElement("UsedOperation", UsedOperation));
            xe.Add(new XElement("UsedKwh", UsedKwh));
            xe.Add(new XElement("UsedCost", UsedCost));
            xe.Add(new XElement("Co2", Co2));
            xe.Add(new XElement("Tree", Tree));
            xe.Add(new XElement("IsSend", IsSend));

            xe.Save(fileName);
        }

        public void LoadFromFile(string fileName = null)
        {
            if (fileName == null) fileName = FileName;
            if (!File.Exists(fileName)) return;

            XElement xe = XElement.Load(fileName);
            
            //_UsedSec = TimeSpan.Parse(xe.Element("UsedSec").Value);
            //_UsedOperation = TimeSpan.Parse(xe.Element("UsedOperation").Value);
            _UsedKwh = double.Parse(xe.Element("UsedKwh").Value);
            _UsedCost = double.Parse(xe.Element("UsedCost").Value);
            _Co2 = double.Parse(xe.Element("Co2").Value);
            _Tree = double.Parse(xe.Element("Tree").Value);
            _IsSend = xe.Element("IsSend").Value;
        }
    }

}

