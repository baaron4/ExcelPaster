using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class PingLoggerObject
    {
        public IPAddress Address { get; }
        public int IntervalTime { get; }
        public string Filepath { get; }

        public PingLoggerObject(IPAddress address, int intervaltime, string filepath)
        {
            Address = address;
            IntervalTime = intervaltime;
            Filepath = filepath;
        }
    }
}
