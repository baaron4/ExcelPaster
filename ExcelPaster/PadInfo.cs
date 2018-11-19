using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class PadInfo
    {

        public PadInfo(string company,string padname, string devicename,string ip,string sub,string gateway)
        {
            this.Company = company;
            this.PadName = padname;
            this.DeviceName = devicename;
            this.IPAddress = ip;
            this.SubnetMask = sub;
            this.Gateway = gateway;

        }
        public string Company;
        public string PadName;
        public string DeviceName;
        public string IPAddress;
        public string SubnetMask;
        public string Gateway;

    }
}
