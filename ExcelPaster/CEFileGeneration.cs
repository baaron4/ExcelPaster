using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    class CEFileGeneration
    {
        List<SiteSystem> SiteSystemList = new List<SiteSystem>();
        int IOHoldingRegister_AppNumber = 9;
        int AnalogInput_ArrayNumber = 0;
        List<PCCUHoldingRegister> AnalogInputs = new List<PCCUHoldingRegister>();
        int DigitalInput_ArrayNumber = 1;
        List<PCCUHoldingRegister> DigitalInputs = new List<PCCUHoldingRegister>();
        int AnalogOutput_ArrayNumber = 2;
        List<PCCUHoldingRegister> AnalogOutputs = new List<PCCUHoldingRegister>();
        int DigitalOutput_ArrayNumber = 3;
        List<PCCUHoldingRegister> DigitalOutputs = new List<PCCUHoldingRegister>();
        int FacilityOperations_AppNumber = 244;
        int Setpoints_ArrayNumber = 208;
        List<PCCUHoldingRegister> Setpoints = new List<PCCUHoldingRegister>();
        //Reuse for 7AM values
        int TanksOperations_AppNumber = 240;
        int SWTanks_ArrayNumber = 200;
        List<PCCUHoldingRegister> SWTanks = new List<PCCUHoldingRegister>();
        int OILTanks_ArrayNumber = 201;
        List<PCCUHoldingRegister> OILTanks = new List<PCCUHoldingRegister>();
        int FWTanks_ArrayNumber = 202;
        List<PCCUHoldingRegister> FWTanks = new List<PCCUHoldingRegister>();
        int Alarm_AppNumber = 94;
        List<PCCUAlarm> AlarmSystem = new List<PCCUAlarm>();
        int AlarmStatus_AppNumber = 242;
        List<PCCUSelect> AlarmStatus_Selects = new List<PCCUSelect>();

        //Helper methods------------------------------------------------------------------------------------------------------------------------
        private string CellValueOrNull(Microsoft.Office.Interop.Excel.Range range, int Y, int X)
        {
            string value = "";
            if (range.Cells[Y, X] != null)
            {
                if (range.Cells[Y, X].Value2 != null)
                {
                    value = range.Cells[Y, X].Value2.ToString();
                }
            }
            return value;
        }
        private string CellDateOrNull(Microsoft.Office.Interop.Excel.Range range, int Y, int X)
        {
            string dvalue = "";
            string value = CellValueOrNull(range, Y, X);
            if (value != "")
            {
                dvalue = DateTime.FromOADate(double.Parse(value)).ToString();
            }
            return dvalue;
        }
        private string FindTitleCell(Microsoft.Office.Interop.Excel.Range range, int Y, int X)
        {
            int SearchUpwardsAmount = 40;
            for (int v = 0; v < SearchUpwardsAmount; v++)
            {
                if (range.Cells[Y - v, X] != null)
                {
                    double thing = range.Cells[Y - v, X].Interior.Color;
                    if (thing != 16777215)
                    {
                        string value = range.Cells[Y - v, X].Value2.ToString();
                        return value.Split('(')[0];
                    }
                }
            }
            return "";
        }
        //Output Methods-----------------------------------------------------------------------------------------------------------------------
        private void SaveHoldingRegisterAsCSV(List<PCCUHoldingRegister> arrayToSave, string fileName, bool isIndirect)
        {
            using (StreamWriter file = new StreamWriter(fileName))
            {
                foreach (PCCUHoldingRegister item in arrayToSave)
                {
                    string line = "";
                    line = item.Description + "," + item.Value + ",";
                    if (isIndirect)
                    {
                        line = line + item.Indirect + ",";
                    }

                    file.WriteLine(line);

                }
            }
        }
        private void SaveAlarmSystemAsCSV(List<PCCUAlarm> arrayToSave, string fileName)
        {
            using (StreamWriter file = new StreamWriter(fileName))
            {
                foreach (PCCUAlarm item in arrayToSave)
                {
                    string line = "";
                    string ThresRegister = item.ThresRegister;
                    if (item.ThresType == "c")
                    {
                        ThresRegister = "";
                    }
                    

                    line = item.Description + ",,y,," + item.Operator + ",,," + item.InputRegister + "," + item.ThresType + "," + ThresRegister + ","
                        + item.ThresConstant + ",,"+item.TriggerType+","+item.TriggerRegister+",,,,," + item.FilterThres + ",," + item.ResetDeadband + ",y,y,n,";
                    file.WriteLine(line);

                }
            }
        }
        //Objects-----------------------------------------------------------------------------------------------------------------------------
        private enum IOType
        { 
            UNKNOWN = 0,DI = 1, AI = 2, DO = 3, AO = 4, RS485 = 5, ETH = 6, VIRTUAL = 7, WIRED = 8
        }
        private enum IOAlarmType
        { 
           NOOP = 0, GT = 1, LT = 2, ON = 3, OFF= 4, AND=5, OR=6, GE=7, LE=8 , NAND=9, NOR=10
        }
        private class IOAlarm
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

            public IOAlarm(string description,IOAlarmType type,bool constantSetpoint,float constantSetpointvalue, float setpoint , float delay)
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
        private class IOPoint
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
        private class Device
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
        private class SiteSystem
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
        private class PCCUAlarm
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
        private class PCCUHoldingRegister
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
            public PCCUHoldingRegister(string register, string description, string value, int cerow)
            {
                this.Register = register;
                this.Description = description;
                this.Value = value;
                this.CERow = cerow;

            }

        }
        private class PCCUSelect
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

        public string GenerateFilesWithExistingCE(string sourceLoc, string outputLoc)
        {
            //Check strings are OK
            if (!File.Exists(sourceLoc))
                return "Existing Cause and Effect File Does not Exist";
            if (!Directory.Exists(outputLoc))
                return "Output Location Needs to be chosen.";
            //Proceed with parseing
            ParseExistingCE(sourceLoc);
            //Generate PCCU Pastes
            GeneratePCCULists();
            //Save CSVs
            SaveHoldingRegisterAsCSV(AnalogInputs,outputLoc + "/IO_HoldingRegisters_AI_Paste.csv",true);
            SaveHoldingRegisterAsCSV(DigitalInputs, outputLoc + "/IO_HoldingRegisters_DI_Paste.csv", true);
            SaveHoldingRegisterAsCSV(AnalogOutputs, outputLoc + "/IO_HoldingRegisters_AO_Paste.csv", true);
            SaveHoldingRegisterAsCSV(DigitalOutputs, outputLoc + "/IO_HoldingRegisters_DO_Paste.csv", true);
            SaveHoldingRegisterAsCSV(Setpoints, outputLoc + "/Setpoints_Paste.csv", false);
            SaveHoldingRegisterAsCSV(SWTanks, outputLoc + "/SWTanks_Paste.csv", true);
            SaveHoldingRegisterAsCSV(OILTanks, outputLoc + "/OilTanks_Paste.csv", true);
            SaveHoldingRegisterAsCSV(FWTanks, outputLoc + "/FWTanks_Paste.csv", true);
            SaveAlarmSystemAsCSV(AlarmSystem,outputLoc + "/SiteAlarmSystem_Paste.csv");

            return "Output files to :" + outputLoc;
        }
       
        private void ParseExistingCEBad(string sourceLoc)
        {
            //Open Excel and read file---------------------------------------------------------------------------------
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sourceLoc);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Cause and Effect"];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            //Aquire colmn desc
            int CEDescriptionColmn = 3;
            int CEFunctionColmn = 4;
            int CERangeColmn = 5;
            int CESetpointColmn = 6;
            int CEUnitsColmn = 7;
            int CESPChangeFromHMI = 8;
            int CEHMIIndicatorColmn = 11;
            int CEHMIAlarmsColmn = 12;

            for (int col = 1; col < xlRange.Columns.Count; col++)
            {
                string value = CellValueOrNull(xlRange, 4, col);
                if (value != "")
                {
                    switch (value.ToUpper())
                    {
                        case "DESCRIPTION":
                            CEDescriptionColmn = col;
                            break;
                        case "FUNCTION":
                            CEFunctionColmn = col;
                            break;
                        case "RANGE":
                            CERangeColmn = col;
                            break;
                        case "SETPOINT":
                            CESetpointColmn = col;
                            break;
                        case "UNITS":
                            CEUnitsColmn = col;
                            break;
                        case "SETPOINT CHANGED FROM HMI":
                            CESPChangeFromHMI = col;
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    value = CellValueOrNull(xlRange, 1, col);
                    if (value != "")
                    {
                        switch (value.ToUpper())
                        {
                            case "HMI INDICATIONS":
                                CEHMIIndicatorColmn = col;
                                break;
                            case "HMI ALARMS":
                                CEHMIAlarmsColmn = col;
                                break;
                            default:
                                break;
                        }
                    }
                }
                
            }
            //generate IO paste

            bool lastrowhadsetpoint = false;

            int SWTankCount = 0;
            int OILTankCount = 0;
            int FWTankCount = 0;
            AnalogInputs.Add(new PCCUHoldingRegister("9.0." + (AnalogInputs.Count), "**Analog Inputs**", "", 0));
            DigitalInputs.Add(new PCCUHoldingRegister("9.2." + (DigitalInputs.Count), "**Digital Inputs**", "", 0));
            AnalogOutputs.Add(new PCCUHoldingRegister("9.1." + (AnalogOutputs.Count), "**Analog Outputs**", "", 0));
            DigitalOutputs.Add(new PCCUHoldingRegister("9.3." + (DigitalOutputs.Count), "**Digital Outputs**", "", 0));
            Setpoints.Add(new PCCUHoldingRegister("244.208." + (Setpoints.Count), "**Setpoints**", "", 0));
            for (int row = 4; row < xlRange.Rows.Count; row++)
            {
                //AI 9.0.X
                if (CellValueOrNull(xlRange, row, CEFunctionColmn) == "AI")
                {

                    string description =FindTitleCell(xlRange, row, CEDescriptionColmn) +" "+ CellValueOrNull(xlRange, row, CEDescriptionColmn);
                    AnalogInputs.Add(new PCCUHoldingRegister("9.0."+ (AnalogInputs.Count), description , "",row));
                }
                //DI 9.2.X
                if (CellValueOrNull(xlRange, row, CEFunctionColmn) == "DI")
                {
                    string description = FindTitleCell(xlRange, row, CEDescriptionColmn) + " " + CellValueOrNull(xlRange, row, CEDescriptionColmn);
                    DigitalInputs.Add(new PCCUHoldingRegister("9.2." + (DigitalInputs.Count), description, "", row));
                }
                //Generate Setpoints paste 244.208.X
                string value = CellValueOrNull(xlRange, row, CESetpointColmn);
                float number;
                if (float.TryParse(value, out number))
                {

                    string description = FindTitleCell(xlRange, row, CEDescriptionColmn) + " " + CellValueOrNull(xlRange, row, CEDescriptionColmn);
                    Setpoints.Add(new PCCUHoldingRegister("244.208." + (Setpoints.Count ), description, CellValueOrNull(xlRange, row, CESetpointColmn), row));
                    lastrowhadsetpoint = true;
                }
                else if (lastrowhadsetpoint)
                {
                    //Create gap between Setpoints
                    Setpoints.Add(new PCCUHoldingRegister("244.208." + (Setpoints.Count + 1), "*", "", 0));
                    lastrowhadsetpoint = false;
                }
                //Generate Tank IO Pastes 240.200.X 240.201.X 240.202.X
                if (CellValueOrNull(xlRange, row, CEFunctionColmn) == "LI")
                {
                    //Look for an interface row
                    int interfaceRow = 0; //0 if no interface
                    int searchforInterface = 0;
                    while (xlRange.Cells[row + searchforInterface, CEFunctionColmn].Interior.Color == 16777215)
                    {
                        if (CellValueOrNull(xlRange, row + searchforInterface, CEFunctionColmn) == "LI")
                        {
                            if (CellValueOrNull(xlRange, row + searchforInterface, CEDescriptionColmn).Contains("INTERFACE"))
                            {
                                interfaceRow = row + searchforInterface;
                                break;
                            }

                        }
                        searchforInterface++;
                    }
                    string description = FindTitleCell(xlRange, row, CEDescriptionColmn);
                    if (description.Contains("SALT"))
                    {
                        SWTankCount++;
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "**" + description + "**", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Status", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " PV (Surface Level) Feet", "", row));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " SV (Interface Level) Feet", "", interfaceRow));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " TV", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " QV", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Comm Response", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Surface FT", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Surface IN", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Interface FT", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "SW" + SWTankCount + " Interface IN", "", 0));
                        SWTanks.Add(new PCCUHoldingRegister("240.200." + (SWTanks.Count), "*", "", 0));
                    }
                    else if (description.Contains("OIL"))
                    {
                        OILTankCount++;
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "**" + description + "**", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Status", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " PV (Surface Level) Feet", "", row));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " SV (Interface Level) Feet", "", interfaceRow));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " TV", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " QV", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Comm Response", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Surface FT", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Surface IN", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Interface FT", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "OIL" + OILTankCount + " Interface IN", "", 0));
                        OILTanks.Add(new PCCUHoldingRegister("240.201." + (OILTanks.Count), "*", "", 0));
                    }
                    else if (description.Contains("FRESH"))
                    {
                        FWTankCount++;
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "**" + description + "**", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Status", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " PV (Surface Level) PSI", "", row));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " SV ", "", interfaceRow));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " TV", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " QV", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Comm Response", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Surface FT", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Surface IN", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Interface FT", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "FW" + FWTankCount + " Interface IN", "", 0));
                        FWTanks.Add(new PCCUHoldingRegister("240.202." + (FWTanks.Count), "*", "", 0));
                    }


                }

                //Generate alarms paste
                if (CellValueOrNull(xlRange, row, CEHMIAlarmsColmn) == "X")
                {
                    //Get alarm name
                    string description = FindTitleCell(xlRange, row, CEDescriptionColmn) + " " + CellValueOrNull(xlRange, row, CEDescriptionColmn);
                    //Get Operation for alarm
                    string function = CellValueOrNull(xlRange, row, CEFunctionColmn).ToUpper();
                    string operation = "";
                    string thresType = "";
                    string thresRegister = "";
                    string thresConstant = "";
                    string setpointvalue ="";
                    int filterthres = 0;
                    int resetthres = 0;
                    int resetdb = 0;
                    if (function.Contains("ALL") || function.Contains("AL"))
                    {
                        operation = "l";
                        thresType = "r";
                        thresRegister = Setpoints.FirstOrDefault(x => x.CERow == row).Register;
                    }
                    else
                    if (function.Contains("AHH") || function.Contains("AH"))
                    {
                        operation = "g";
                        thresType = "r";
                        thresRegister = Setpoints.FirstOrDefault(x => x.CERow == row).Register;
                    }
                    else if (function.Contains("FAULT"))
                    {
                         setpointvalue = CellValueOrNull(xlRange, row, CESetpointColmn);
                        if (setpointvalue == "COMM")
                        {
                            operation = "g";
                            thresType = "c";
                            thresConstant = "0";
                        }
                        else if (setpointvalue == "NO")
                        {
                            operation = "l";
                            thresType = "c";
                            thresConstant = "1";
                        }else
                        {
                            operation = "l";
                            thresType = "r";
                            PCCUHoldingRegister sp = Setpoints.FirstOrDefault(x => x.CERow == row);
                            if (sp != null)
                            {
                                thresRegister = sp.Register;
                            }
                            
                        }
                    }
                    else if (function.Contains("DI"))
                    {
                         setpointvalue = CellValueOrNull(xlRange, row, CESetpointColmn);
                        if (setpointvalue == "NC")
                        {
                            operation = "l";
                            thresType = "c";
                            thresConstant = "1";
                        }
                    }
                    //Get alarm input register
                    //Look for an AI row
                    int airow = 0;
                    int dirow = 0;
                    if (setpointvalue != "COMM")
                    {
                        int searchforAI = 0;
                        double color = xlRange.Cells[row - searchforAI, CEFunctionColmn].Interior.Color;
                        while (color == 16777215)
                        {
                            if (CellValueOrNull(xlRange, row - searchforAI, CEFunctionColmn).Contains("AI"))
                            {
                                airow = row - searchforAI;
                                break;
                            }
                            else if (CellValueOrNull(xlRange, row - searchforAI, CEFunctionColmn).Contains("DI"))
                            {
                                dirow = row - searchforAI;
                                break;
                            }
                            if (searchforAI > xlRange.Rows.Count)
                            {
                                break;
                            }
                            color = xlRange.Cells[row - searchforAI, CEFunctionColmn].Interior.Color;
                            searchforAI++;
                        }
                    }
                   
                    string inputRegister = "";
                    if (airow != 0)
                    {
                        PCCUHoldingRegister aiHR = AnalogInputs.FirstOrDefault(x => x.CERow == airow);
                        if (aiHR != null)
                        {
                            inputRegister = aiHR.Register;
                        }
                    }
                    else if (dirow != 0)
                    {
                        PCCUHoldingRegister diHR = DigitalInputs.FirstOrDefault(x => x.CERow == dirow);
                        if (diHR != null)
                        {
                            inputRegister = diHR.Register;
                        }
                    }

                    //Specefeic alarms
                    if (description.ToUpper().Contains("FAIL TO START"))
                    {
                        filterthres = 30;
                    }
                    if (function.ToUpper().Contains("FAULT"))
                    {
                        filterthres = 60;
                    }
                    AlarmSystem.Add(new PCCUAlarm("94.122." + (AlarmSystem.Count), description, operation, inputRegister, thresType, thresRegister, thresConstant, filterthres, resetdb));
                }
            }

            for (int col = 15; col < xlRange.Columns.Count; col++)
            {
                //AO 9.1.X
                if (CellValueOrNull(xlRange, 3, col) == "AO")
                {
                    AnalogOutputs.Add(new PCCUHoldingRegister("9.1." + (AnalogOutputs.Count ), CellValueOrNull(xlRange, 1, col), "", col));
                }
                //DO 9.3.X
                if (CellValueOrNull(xlRange, 3, col) == "DO")
                {
                    DigitalOutputs.Add(new PCCUHoldingRegister("9.3." + (DigitalOutputs.Count), CellValueOrNull(xlRange, 1, col), "", col));
                }
            }

            //Generate Select for alarms
            int selectCount = 0;
            int wordCount = 1;
            int bitCount = 0;
            foreach (PCCUAlarm alarm in AlarmSystem)
            {
                AlarmStatus_Selects.Add(new PCCUSelect("242.32." + selectCount,wordCount + "." + bitCount,alarm.Register,"","242.201."+ bitCount));
                selectCount++;
                bitCount++;
                if (selectCount % 16 == 0)
                {
                    wordCount++;
                    bitCount = 0;
                }
            }
            xlWorkbook.Close();
            xlApp.Quit();
            
        }

        private void ParseExistingCE(string sourceLoc)
        {
            //Open Excel and read file---------------------------------------------------------------------------------
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sourceLoc);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Cause and Effect"];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            //Aquire column desc
            int CEPIDColumn = 1;
            int CEDescriptionColmn = 3;
            int CEFunctionColmn = 4;
            int CERangeColmn = 5;
            int CESetpointColmn = 6;
            int CEUnitsColmn = 7;
            int CESPChangeFromHMI = 8;
            int CEHMIIndicatorColmn = 11;
            int CEHMIAlarmsColmn = 12;

            for (int col = 1; col < xlRange.Columns.Count; col++)
            {
                string value = CellValueOrNull(xlRange, 4, col);
                if (value != "")
                {
                    switch (value.ToUpper())
                    {
                        case "P&ID":
                            CEPIDColumn = col;
                            break;
                        case "DESCRIPTION":
                            CEDescriptionColmn = col;
                            break;
                        case "FUNCTION":
                            CEFunctionColmn = col;
                            break;
                        case "RANGE":
                            CERangeColmn = col;
                            break;
                        case "SETPOINT":
                            CESetpointColmn = col;
                            break;
                        case "UNITS":
                            CEUnitsColmn = col;
                            break;
                        case "SETPOINT CHANGED FROM HMI":
                            CESPChangeFromHMI = col;
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    value = CellValueOrNull(xlRange, 1, col);
                    if (value != "")
                    {
                        switch (value.ToUpper())
                        {
                            case "HMI INDICATIONS":
                                CEHMIIndicatorColmn = col;
                                break;
                            case "HMI ALARMS":
                                CEHMIAlarmsColmn = col;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }
            //Go Over all rows
            int firstRowofDevice = 0;
            int amountOfMergedRows = 0;
            for (int row = 4; row < xlRange.Rows.Count; row++)
            {
                double cellColor = xlRange.Cells[row, CEDescriptionColmn].Interior.Color;
                string cellDesc = CellValueOrNull(xlRange, row, CEDescriptionColmn);
                
                if (cellColor != 16777215)//If Title Row
                {
                    SiteSystemList.Add(new SiteSystem(xlRange.Cells[row, CEDescriptionColmn].Value2(), SiteSystemList.Count +1));
                }
                else 
                {//must be device value
                    string pidName = CellValueOrNull(xlRange,row, CEPIDColumn);
                    string function = CellValueOrNull(xlRange, row, CEFunctionColmn);
                    string unit = CellValueOrNull(xlRange, row, CEUnitsColmn);
                    string HMIAlarm = CellValueOrNull(xlRange, row, CEHMIAlarmsColmn);
                    string[] range = CellValueOrNull(xlRange, row, CERangeColmn).Split('|');
                    float LRV = 0; 
                    float URV = 0;
                    int alarmDelay = 0;
                    if (range != null && range[0] !="")
                    {
                        LRV = float.Parse(range[0].TrimStart('(').TrimEnd(')'));
                        URV = float.Parse(range[1].TrimStart('(').TrimEnd(')').TrimEnd('?'));
                    }
                    string cellsetpoint = CellValueOrNull(xlRange, row, CESetpointColmn);
                    bool constantSetpoint = false;
                    int constantSetpointValue = 0;


                    Microsoft.Office.Interop.Excel.Range mergerange = null;
                    mergerange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[row, CEPIDColumn];

                    
                    object mergedCells = mergerange.MergeCells;
                   
                    if ((bool)mergedCells)
                    {
                        
                        amountOfMergedRows = mergerange.MergeArea.Rows.Count;

                    }
                   
                    
                    if ((bool)mergedCells && firstRowofDevice == 0)//First Line of larger device
                    {
                        Device newDevice = new Device(cellDesc, pidName);
                        IOPoint newPoint = new IOPoint(cellDesc, function, unit, LRV, URV);


                        if (HMIAlarm == "X" )
                        {
                            IOAlarmType operationType = IOAlarmType.NOOP;
                            if (function.Contains("AH") || function.Contains("AHH"))
                            {
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("AL") || function.Contains("ALL"))
                            {
                                operationType = IOAlarmType.LT;
                            } else if (function.Contains("FAULT") && cellsetpoint.Contains("COMM"))
                            {
                                operationType = IOAlarmType.GT;
                                alarmDelay = 60;
                            } else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NC"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 1;
                                operationType = IOAlarmType.LT;
                            }
                            else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NC"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 0;
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("FAULT") )
                            {
                               
                                operationType = IOAlarmType.LT;
                            }
                            float setpoint = 0;
                            if (float.TryParse(cellsetpoint, out setpoint))
                            {
                                //Nothing to do here now?
                            }
                           
                            newPoint.alarmList.Add(new IOAlarm(cellDesc,operationType,constantSetpoint, constantSetpointValue, setpoint, alarmDelay));
                            
                        }

                        newDevice.IOPointList.Add(newPoint);
                        SiteSystemList.Last().DeviceList.Add(newDevice);
                        firstRowofDevice = row;
                        
                    }
                    else if ((bool)mergedCells && firstRowofDevice != 0)//multline device, additional IO OR is a complete device with multiple lines
                    {
                        if (HMIAlarm == "X")//if alarm
                        {
                            IOAlarmType operationType = IOAlarmType.NOOP;
                            if (function.Contains("AH") || function.Contains("AHH"))
                            {
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("AL") || function.Contains("ALL"))
                            {
                                operationType = IOAlarmType.LT;
                            }
                            else if (function.Contains("FAULT") && cellsetpoint.Contains("COMM"))
                            {
                                operationType = IOAlarmType.GT;
                                alarmDelay = 60;
                            }
                            else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NC"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 1;
                                operationType = IOAlarmType.LT;
                            }
                            else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NO"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 0;
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("FAULT"))
                            {

                                operationType = IOAlarmType.LT;
                            }
                            float setpoint = 0;
                            if (float.TryParse(cellsetpoint, out setpoint))
                            {
                                //Nothing to do here now?
                            }
                            SiteSystemList.Last().DeviceList.Last().IOPointList.Last().alarmList.Add(new IOAlarm(cellDesc, operationType, constantSetpoint, constantSetpointValue,setpoint, alarmDelay));
                        }
                        else 
                        {
                            IOPoint newPoint = new IOPoint(cellDesc, function, unit, LRV, URV);
                            SiteSystemList.Last().DeviceList.Last().IOPointList.Add(newPoint);
                        }
                        if (row+1-firstRowofDevice == amountOfMergedRows)
                        {
                            firstRowofDevice = 0;
                            amountOfMergedRows = 0;
                        }
                        
                    }
                    else if(!(bool)mergedCells)//is a 1 line device
                    {
                        Device newDevice = new Device(cellDesc, pidName);
                        IOPoint newPoint = new IOPoint(cellDesc, function, unit, LRV, URV);
                        float setpoint = 0;
                        if( float.TryParse(cellsetpoint, out setpoint))
                        { 
                            //Nothing to do here now?
                        }
                       

                        if (HMIAlarm == "X")
                        {
                            IOAlarmType operationType = IOAlarmType.NOOP;
                            if (function.Contains("AH") || function.Contains("AHH"))
                            {
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("AL") || function.Contains("ALL"))
                            {
                                operationType = IOAlarmType.LT;
                            }
                            else if (function.Contains("FAULT") && cellsetpoint.Contains("COMM"))
                            {
                                operationType = IOAlarmType.GT;
                                alarmDelay = 60;
                            }
                            else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NC"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 1;
                                operationType = IOAlarmType.LT;
                            }
                            else if ((function.Contains("FAULT") || function.Contains("DI")) && cellsetpoint.Contains("NO"))
                            {
                                constantSetpoint = true;
                                constantSetpointValue = 0;
                                operationType = IOAlarmType.GT;
                            }
                            else if (function.Contains("FAULT"))
                            {

                                operationType = IOAlarmType.LT;
                            }
                            newPoint.alarmList.Add(new IOAlarm(cellDesc, operationType, constantSetpoint, constantSetpointValue, setpoint, alarmDelay));
                        }
                        newDevice.IOPointList.Add(newPoint);
                        SiteSystemList.Last().DeviceList.Add(newDevice);

                    }
                }

            }

            //Go over all output columns
            for (int col = 15; col < xlRange.Columns.Count; col++)
            {
                string cellDescription = CellValueOrNull(xlRange, 1, col);
                string cellPID = CellValueOrNull(xlRange, 2, col);
                string cellType = CellValueOrNull(xlRange, 3, col);

                if (cellPID != "")
                {
                    foreach (SiteSystem system in SiteSystemList)
                    {
                        //Device device = system.DeviceList.FirstOrDefault(x => x.PID == cellPID);
                        foreach (Device device in system.DeviceList)
                        {
                            if (device != null)
                            {
                                if (device.PID == cellPID)
                                {
                                    device.IOPointList.Add(new IOPoint(cellDescription, cellType, "Output", 0, 0));
                                    //Check for any Inputs that need a output to activate
                                    foreach (IOPoint p in device.IOPointList)
                                    {
                                        foreach (IOAlarm a in p.alarmList)
                                        {
                                            if (a.Description.Contains("FAIL TO START"))
                                            {
                                                a.triggerType = "r";
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            xlWorkbook.Close();
            xlApp.Quit();
            
        }

        private void GeneratePCCULists()
        {
            //Go over all systems and create PCCU Registers------------------------------------------------------------
            //Setup array titles
            AnalogInputs.Add(new PCCUHoldingRegister(IOHoldingRegister_AppNumber + "." + AnalogInput_ArrayNumber + "." + (AnalogInputs.Count), "**Analog Inputs**", "", 0));
            DigitalInputs.Add(new PCCUHoldingRegister(IOHoldingRegister_AppNumber + "." + DigitalInput_ArrayNumber + "." + (DigitalInputs.Count), "**Digital Inputs**", "", 0));
            AnalogOutputs.Add(new PCCUHoldingRegister(IOHoldingRegister_AppNumber + "." + AnalogOutput_ArrayNumber + "." + (AnalogOutputs.Count), "**Analog Outputs**", "", 0));
            DigitalOutputs.Add(new PCCUHoldingRegister(IOHoldingRegister_AppNumber + "." + DigitalOutput_ArrayNumber + "." + (DigitalOutputs.Count), "**Digital Outputs**", "", 0));
            Setpoints.Add(new PCCUHoldingRegister(FacilityOperations_AppNumber + "." + Setpoints_ArrayNumber + "." + (Setpoints.Count), "**Setpoints**", "", 0));

            int selectCount = 0;
            int wordCount = 1;
            int bitCount = 0;
            int SWTankCount = 0;
            int OILTankCount = 0;
            int FWTankCount = 0;
            foreach (SiteSystem system in SiteSystemList)
            {
                foreach (Device device in system.DeviceList)
                {
                    //Setpoints
                    Setpoints.Add(new PCCUHoldingRegister(FacilityOperations_AppNumber + "." + Setpoints_ArrayNumber + "." + (Setpoints.Count), "**" + device.Name + "**", "", 0));
                    //Tanks
                    if (device.PID.Contains("LIT"))
                    {
                        if (system.Name.Contains("FRESH WATER"))
                        {
                            FWTankCount++;
                            string registerArray = TanksOperations_AppNumber + "." + FWTanks_ArrayNumber + ".";
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "**" + device.PID + "**", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Status", ""));

                            string register = registerArray + (FWTanks.Count);
                            FWTanks.Add(new PCCUHoldingRegister(register, "FW" + FWTankCount + " PV (Surface Level) PSI", ""));
                            IOPoint surfPoint = device.IOPointList.First(x => x.Description.Contains("SURFACE"));
                            if (surfPoint != null)
                            {
                                device.IOPointList.First(x => x.Description.Contains("SURFACE")).PCCUHoldingRegister = register;
                            }
                            
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " SV ", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " TV", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " QV", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Comm Response", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Surface FT", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Surface IN", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Interface FT", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "FW" + FWTankCount + " Interface IN", ""));
                            FWTanks.Add(new PCCUHoldingRegister(registerArray + (FWTanks.Count), "*", "", 0));
                           
                        }
                        else if (system.Name.Contains("OIL"))
                        {
                            OILTankCount++;
                            string registerArray = TanksOperations_AppNumber + "." + OILTanks_ArrayNumber + ".";
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "**" + device.PID + "**", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Status", ""));

                            string register = registerArray + (OILTanks.Count);
                            OILTanks.Add(new PCCUHoldingRegister(register, "OIL" + OILTankCount + " PV (Surface Level) Feet", ""));
                            IOPoint surfPoint = device.IOPointList.First(x => x.Description.Contains("SURFACE"));
                            if (surfPoint != null)
                            {
                                device.IOPointList.First(x => x.Description.Contains("SURFACE")).PCCUHoldingRegister = register;
                            }

                            register = registerArray + (OILTanks.Count);
                            OILTanks.Add(new PCCUHoldingRegister(register, "OIL" + OILTankCount + " SV (Interface Level) Feet", ""));
                            IOPoint intPoint = device.IOPointList.First(x => x.Description.Contains("INTERFACE"));
                            if (intPoint != null)
                            {
                                device.IOPointList.First(x => x.Description.Contains("INTERFACE")).PCCUHoldingRegister = register;
                            }

                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " TV", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " QV", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Comm Response", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Surface FT", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Surface IN", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Interface FT", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "OIL" + OILTankCount + " Interface IN", ""));
                            OILTanks.Add(new PCCUHoldingRegister(registerArray + (OILTanks.Count), "*", "", 0));
                        }
                        else if (system.Name.Contains("SALT WATER"))
                        {
                            SWTankCount++;
                            string registerArray = TanksOperations_AppNumber + "." + SWTanks_ArrayNumber + ".";
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "**" + device.PID + "**", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Status", ""));

                            string register = registerArray + (SWTanks.Count);
                            SWTanks.Add(new PCCUHoldingRegister(register, "SW" + SWTankCount + " PV (Surface Level) Feet", ""));
                            IOPoint surfPoint = device.IOPointList.First(x => x.Description.Contains("SURFACE"));
                            if (surfPoint != null)
                            {
                                device.IOPointList.First(x => x.Description.Contains("SURFACE")).PCCUHoldingRegister = register;
                            }

                            register = registerArray + (SWTanks.Count);
                            SWTanks.Add(new PCCUHoldingRegister(register, "SW" + SWTankCount + " SV (Interface Level) Feet", ""));
                            IOPoint intPoint = device.IOPointList.First(x => x.Description.Contains("INTERFACE"));
                            if (intPoint != null)
                            {
                                device.IOPointList.First(x => x.Description.Contains("INTERFACE")).PCCUHoldingRegister = register;
                            }

                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " TV", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " QV", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Comm Response", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Surface FT", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Surface IN", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Interface FT", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "SW" + SWTankCount + " Interface IN", ""));
                            SWTanks.Add(new PCCUHoldingRegister(registerArray + (SWTanks.Count), "*", ""));
                        }
                    }
                    foreach (IOPoint point in device.IOPointList)
                    {
                        string pointName = system.Name +/* " " + device.Name +*/ " " + point.Description;
                        string PCCUHoldingRegister = "";
                        if (point.type == IOType.AI)
                        {
                            PCCUHoldingRegister = IOHoldingRegister_AppNumber + "." + AnalogInput_ArrayNumber + "." + (AnalogInputs.Count);
                            AnalogInputs.Add(new PCCUHoldingRegister(PCCUHoldingRegister, pointName, ""));
                        }
                        else if (point.type == IOType.DI)
                        {
                            PCCUHoldingRegister = IOHoldingRegister_AppNumber + "." + DigitalInput_ArrayNumber + "." + (DigitalInputs.Count);
                            DigitalInputs.Add(new PCCUHoldingRegister(PCCUHoldingRegister, pointName, ""));
                        }
                        else if (point.type == IOType.AO)
                        {
                            PCCUHoldingRegister = IOHoldingRegister_AppNumber + "." + AnalogOutput_ArrayNumber + "." + (AnalogOutputs.Count);
                            AnalogOutputs.Add(new PCCUHoldingRegister(PCCUHoldingRegister, pointName, ""));
                        }
                        else if (point.type == IOType.DO)
                        {
                            PCCUHoldingRegister = IOHoldingRegister_AppNumber + "." + DigitalOutput_ArrayNumber + "." + (DigitalOutputs.Count);
                            DigitalOutputs.Add(new PCCUHoldingRegister(PCCUHoldingRegister, pointName, ""));
                            //Check for fail to start alarms
                            foreach (IOPoint p in device.IOPointList)
                            {
                                if (p.alarmList != null)
                                {
                                    //IOAlarm checkAlarm = p.alarmList.FirstOrDefault(x => x.Description.Contains("FAIL TO START"));
                                    foreach (IOAlarm a in p.alarmList)
                                    {
                                        if (a.Description.Contains("FAIL TO START"))
                                        {
                                           
                                                a.triggerRegister = PCCUHoldingRegister;
                                                break;
                                            
                                        }
                                    }
                                }
                            }
                        }
                        //TODO: add Coriolis and Mag meters
                        point.PCCUHoldingRegister = PCCUHoldingRegister;
                        //Alarm Lists
                        foreach (IOAlarm alarm in point.alarmList)
                        {
                            //PCCU Setpoints Holding Register
                            string setpointRegister = "";
                            if (alarm.setpoint != null)
                            {
                                if (!alarm.constantSetpoint)
                                {
                                    setpointRegister = FacilityOperations_AppNumber + "." + Setpoints_ArrayNumber + "." + (Setpoints.Count);
                                    Setpoints.Add(new PCCUHoldingRegister(setpointRegister, alarm.Description, alarm.setpoint.ToString()));
                                    alarm.PCCUSetpointRegister = setpointRegister;
                                }
                                
                            }
                            //PCCU alarm System
                            string thresType = "r";
                            if (alarm.constantSetpoint)
                            {
                                thresType = "c";
                            }
                            string PCCUAlarmRegister = Alarm_AppNumber + ".122." + AlarmSystem.Count;
                            AlarmSystem.Add(new PCCUAlarm(PCCUAlarmRegister, system.Name + " " + alarm.Description, alarm.IOAlarmToString(), point.PCCUHoldingRegister,
                                thresType, setpointRegister, alarm.constantSetpointValue.ToString(), (int)alarm.delay, (int)alarm.resetdelay));
                            alarm.PCCUAlarmRegister = PCCUAlarmRegister;

                            AlarmStatus_Selects.Add(new PCCUSelect(AlarmStatus_AppNumber + ".32." + selectCount, wordCount + "." + bitCount, alarm.PCCUAlarmRegister, "", AlarmStatus_AppNumber +".201." + bitCount));
                            selectCount++;
                            bitCount++;
                            if (selectCount % 16 == 0)
                            {
                                wordCount++;
                                bitCount = 0;
                            }
                        }
                    }
                }
            }
        }
    }
}
