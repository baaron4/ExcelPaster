using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class PCCUObjects
    {
       
    }
    public class PCCUAlarm
    {
        public string Register;
        public string Description;
        public string Operator;
        public string InputRegister;
        public string ThresType;
        public string ThresRegister;
        public string ThresConstant;
        public string TriggerType;
        public string TriggerRegister;
        public int FilterThres;
        public int ResetDeadband;
        public PCCUAlarm(string register, string desc, string op, string input, string thresType, string thresReg, string thresConstant, int filterThres, int resetDB)
        {
            this.Register = register;
            this.Description = desc;
            this.Operator = op;
            this.InputRegister = input;
            this.ThresType = thresType;
            this.ThresRegister = thresReg;
            this.ThresConstant = thresConstant;
            this.FilterThres = filterThres;
            this.ResetDeadband = resetDB;
        }
    }
    public class PCCUHoldingRegister
    {
        public string Register;
        public string Description;
        public string Value;
        public string Indirect;
        public int CERow;

        public PCCUHoldingRegister(string register, string description, string value)
        {
            this.Register = register;
            this.Description = description;
            this.Value = value;

        }


    }
    public class PCCUSelect
    {
        public string Register;
        public string Description;
        public string SelectRegister;
        public string Register1;
        public string Register2;

        public PCCUSelect(string register, string desc, string selectregister, string reg1, string reg2)
        {
            this.Register = register;
            this.Description = desc;
            this.SelectRegister = selectregister;
            this.Register1 = reg1;
            this.Register2 = reg2;

        }
    }
}
