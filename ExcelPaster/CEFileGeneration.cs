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

        List<PCCUHoldingRegister> AnalogInputs = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> DigitalInputs = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> AnalogOutputs = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> DigitalOutputs = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> Setpoints = new List<PCCUHoldingRegister>();
        //Reuse for 7AM values
        List<PCCUHoldingRegister> SWTanks = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> OILTanks = new List<PCCUHoldingRegister>();
        List<PCCUHoldingRegister> FWTanks = new List<PCCUHoldingRegister>();

        List<PCCUAlarm> AlarmSystem = new List<PCCUAlarm>();

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
                    line = item.Description + ",,y,," + item.Operator + ",,," + item.InputRegister + "," + item.ThresType + "," + item.ThresRegister + ","
                        + item.ThresConstant + ",,,,,,,," + item.FilterThres + ",," + item.ResetDeadband + ",y,y,n,";
                    file.WriteLine(line);

                }
            }
        }
        //Objects-----------------------------------------------------------------------------------------------------------------------------
        private enum IOType
        { 
            DI = 1, AI = 2, DO = 3, AO = 4, RS485 = 5, ETH = 6
        }
        private enum IOAlarmType
        { 
           NOOP = 0, GT = 1, LT = 2, ON = 3, OFF= 4, AND=5, OR=6, GE=7, LE=8 , NAND=9, NOR=10
        }
        private class IOAlarm
        {
            public string Description;
            public IOAlarmType type;
            public float setpoint;
            public float delay;
            public float resetdelay;
        }
        private class IOPoint
        {
            public string Description;
            public IOType type;
            public string unit;
            public float LRV;
            public float URV;
            public List<IOAlarm> alarmList;

        }
        private class Device
        {
            public string Name;
            public string PID;
            public List<IOPoint> IOPointList;

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
            ParseExistingCEBad(sourceLoc);
            //Generate PCCU Pastes
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
            for (int row = 4; row < xlRange.Rows.Count; row++)
            {
                int cellColor = xlRange.Cells[row, CEDescriptionColmn].Interior.Color;
                string cellDesc = CellValueOrNull(xlRange, row, CEDescriptionColmn);
                
                if (cellColor != 16777215)//If Title Row
                {
                    SiteSystemList.Add(new SiteSystem(xlRange.Cells[row, CEDescriptionColmn].Value2(), SiteSystemList.Count +1));
                }
                else 
                {

                    string cellFunction = CellValueOrNull(xlRange, row, CEFunctionColmn);
                    
                    //if(cellFunction == )

                }

            }
        }
    }
}
