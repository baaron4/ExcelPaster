using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class SystemObjects
    {
    }

    public enum IOType
    {
        UNKNOWN = 0, DI = 1, AI = 2, DO = 3, AO = 4, RS485 = 5, ETH = 6, VIRTUAL = 7, WIRED = 8
    }
    public enum IOAlarmType
    {
        NOOP = 0, GT = 1, LT = 2, ON = 3, OFF = 4, AND = 5, OR = 6, GE = 7, LE = 8, NAND = 9, NOR = 10
    }
    public class IOAlarm
    {
        public string Description;
        public IOAlarmType type;
        public bool constantSetpoint;
        public float constantSetpointValue;
        public float setpoint;
        public float delay;
        public float resetdelay;
        public string triggerType;
        public string triggerRegister;
        public string PCCUAlarmRegister;
        public string PCCUSetpointRegister;

        public IOAlarm(string description, IOAlarmType type, bool constantSetpoint, float constantSetpointvalue, float setpoint, float delay)
        {
            this.Description = description;
            this.type = type;
            this.constantSetpoint = constantSetpoint;
            this.constantSetpointValue = constantSetpointvalue;
            this.setpoint = setpoint;
            this.delay = delay;
        }

        public string IOAlarmToString()
        {
            switch (this.type)
            {
                case IOAlarmType.GT:
                    return "g";
                case IOAlarmType.LT:
                    return "l";
                case IOAlarmType.GE:
                    return "gg";
                case IOAlarmType.LE:
                    return "ll";
                case IOAlarmType.AND:
                    return "a";
                case IOAlarmType.OR:
                    return "o";
                default:
                    return "";
            }
        }
    }
    public class IOPoint
    {
        public string Description;
        public IOType type;
        public string unit;
        public float LRV;
        public float URV;
        public List<IOAlarm> alarmList;
        public string PCCUHoldingRegister;
        public IOPoint(string Description, string type, string unit, float LRV, float URV)
        {
            IOType pointtype = IOType.UNKNOWN;
            switch (type)
            {
                case "DI":
                    pointtype = IOType.DI;
                    break;
                case "AI":
                    pointtype = IOType.AI;
                    break;
                case "DO":
                    pointtype = IOType.DO;
                    break;
                case "AO":
                    pointtype = IOType.AO;
                    break;
                case "VAI":
                    pointtype = IOType.VIRTUAL;
                    break;
                case "RELAY":
                    pointtype = IOType.WIRED;
                    break;
            }
            this.Description = Description;
            this.type = pointtype;
            this.unit = unit;
            this.LRV = LRV;
            this.URV = URV;
            this.alarmList = new List<IOAlarm>();
        }
    }
    public class Device
    {
        public string Name;
        public string PID;
        public List<IOPoint> IOPointList;
        public Device(string name, string PID)
        {
            this.Name = name;
            this.PID = PID;
            this.IOPointList = new List<IOPoint>();
        }
    }
    public class SiteSystem
    {
        public string Name;
        public int Number;
        public List<Device> DeviceList;

        public SiteSystem(string name, int number)
        {
            this.Name = name;
            this.Number = number;
            DeviceList = new List<Device>();
        }
    }
}
