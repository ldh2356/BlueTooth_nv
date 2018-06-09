using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSES_Program
{
    class SendPCEnergy
    {
        public string eventType { get; set; }
        public string macAddress { get; set; }
        public string uptime { get; set; }
        public string savingTime { get; set; }

        public Dictionary<string, string> hardwardInfo { get; set; }
    }
}
