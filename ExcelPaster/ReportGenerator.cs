using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using iText.Forms;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelPaster
{
    class ReportGenerator
    {
        public string printDateTime = "", analyzedBy = "", meterID = "", meterDescription = "", analysisTime = "", sampleType = "", elevation = "";
        public float flowingTemp = 0, flowingPressure = 0, calibrationElevation = 0,
            locationElevation = 0, inferiorWobbe = 0, superiorWobbe = 0,
            compressibility = 0, density = 0, realRelDensity = 0, idealCV = 0, wetCV = 0, dryCV = 0, contractTemp = 0, contractPress = 0, atmoPressure = 0;
        public int numCycles = 0, connectedStreams = 0;

        public static readonly String FONT = System.AppDomain.CurrentDomain.BaseDirectory + "Resources\\FreeSans.ttf";
        private float GetNumbersAndDecimalsAsFloat(string input)
        {
            string st = new string(input.Where(c => char.IsDigit(c) || c == '.').ToArray());

            return float.Parse(st);
        }
        void DrawImage(XGraphics gfx, string jpegSamplePath, int x, int y, int width, int height)
        {
            XImage image = XImage.FromFile(jpegSamplePath);
            gfx.DrawImage(image, x, y, width, height);
        }
        public class Gas
        {
            public string Name;
            public float UnNorm;
            public float Norm;
            public float Liquids;
            public float Ideal;
            public float RelDensity;

            public Gas(string name, float unNorm, float norm, float liquids, float ideal, float relDensity)
            {
                this.Name = name;
                this.UnNorm = unNorm;
                this.Norm = norm;
                this.Liquids = liquids;
                this.Ideal = ideal;
                this.RelDensity = relDensity;
            }
        }
        //Make data-fetching code its own function
        public List<Gas> loadData(string sourceLoc)
        {
            List<Gas> gasList = new List<Gas>();

            //Read text file
            // Read a text file line by line.  
            string[] lines = File.ReadAllLines(sourceLoc);
            int lineNum = 0;
            bool tablePrimer = false;
            foreach (string line in lines)
            {
                switch (lineNum)
                {
                    case 0:
                        if (line.Contains("Print Date Time:"))
                        {
                            printDateTime = line.Replace("Print Date Time:", "").Replace("  ", "");
                            lineNum++;
                        }
                        break;
                    case 1:
                        if (line.Contains("Analyzed By:"))
                        {
                            analyzedBy = line.Replace("Analyzed By:", "").Replace("  ", "");
                            lineNum++;
                        }
                        break;
                    case 2:
                        if (line.Contains("Meter ID:"))
                        {
                            meterID = line.Replace("Meter ID:", "").Replace("  ", "").TrimEnd('.');
                            lineNum++;
                        }
                        break;
                    case 3:
                        if (line.Contains("..."))
                        {
                            meterDescription = line.Trim().Split('.')[0];
                        }
                        if (line.Contains("Analysis Time:"))
                        {
                            analysisTime = line.Substring(0, line.LastIndexOf("Sample Type:")).Replace("Analysis Time:", "").Replace("  ", "");
                        }
                        if (line.Contains("Sample Type:"))
                        {
                            sampleType = line.Substring(line.LastIndexOf("Sample Type:"), line.Length - line.LastIndexOf("Sample Type:")).Replace("  ", "").Replace("Sample Type:", "");
                            lineNum++;
                        }
                        break;
                    case 4:
                        if (line.Contains("Flowing Temp.:"))
                        {
                            if (line.Contains("-")) line.Replace("-", "0");
                            if (line.Contains("?")) line.Replace("?", "0");
                            if(line.Substring(0, line.LastIndexOf("Flowing Pressure:")).Replace("Flowing Temp.:", "").Replace("  ", "").Replace("Deg. F", "") == "")
                            {
                                flowingTemp = 0;
                            }
                            else flowingTemp = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Flowing Pressure:")).Replace("Flowing Temp.:", "").Replace("  ", "").Replace("Deg. F", ""));
                        }
                        if (line.Contains("Flowing Pressure:"))
                        {
                            if (line.Contains("-")) line.Replace("-", "0");
                            if (line.Contains("?")) line.Replace("?", "0");
                            if(line.Substring(line.LastIndexOf("Flowing Pressure:"), line.Length - line.LastIndexOf("Flowing Pressure:")).Replace("  ", "") == "")
                            {
                                flowingPressure = 0;
                            }
                            else flowingPressure = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Flowing Pressure:"), line.Length - line.LastIndexOf("Flowing Pressure:")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 5:
                        if (line.Contains("Calibration Elevation:"))
                        {
                            calibrationElevation = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Location Elevation:")).Replace("Calibration Elevation:", "").Replace("  ", ""));
                        }
                        else
                        {
                            lineNum++;
                        }
                        if (line.Contains("Location Elevation:"))
                        {
                            locationElevation = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Location Elevation:"), line.Length - line.LastIndexOf("Location Elevation:")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 6://into table
                        if (tablePrimer == false)
                        {
                            if (line.Contains("----------------------------------------------------------------------------")) tablePrimer = true;
                        }
                        else
                        {
                            string combineWords = line.Replace("Carbon Dioxide", "Carbon-Dioxide");

                            string reduceSpaces = string.Join(" ", combineWords.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
                            combineWords = reduceSpaces.Replace("Hydrogen Sulfide", "Hydrogen-Sulfide");
                            reduceSpaces = string.Join(" ", combineWords.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));

                            string[] data = reduceSpaces.Split(' ');
                            if (data.Length >= 5)
                            {
                                gasList.Add(new Gas(data[0], float.Parse(data[1]), float.Parse(data[2]),
                                    float.Parse(data[3]), float.Parse(data[4]), float.Parse(data[5])));
                            }
                            if (data[0] == "Total")
                            {
                                
                                lineNum++;
                            }
                        }
                        break;
                    case 7:
                        if (line.Contains("Elevation"))
                        {
                            elevation = line.Replace("Elevation", "").Replace("  ", "");
                            lineNum++;
                        }
                        else
                        {
                            lineNum++;
                        }
                        break;
                    case 8:
                        if (line.Contains("Inferior Wobbe"))
                        {
                            inferiorWobbe = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Superior Wobbe")).Replace("Inferior Wobbe", "").Replace("  ", ""));
                        }
                        if (line.Contains("Superior Wobbe"))
                        {
                            superiorWobbe = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Superior Wobbe"), line.Length - line.LastIndexOf("Superior Wobbe")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 9:
                        if (line.Contains("Compressibility"))
                        {
                            compressibility = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Density")).Replace("Compressibility", "").Replace("  ", ""));
                        }
                        if (line.Contains("Density"))
                        {
                            density = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Density"), line.Length - line.LastIndexOf("Density")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 10:
                        if (line.Contains("Real Rel. Density"))
                        {
                            realRelDensity = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Ideal CV")).Replace("Real Rel. Density", "").Replace("  ", ""));
                        }
                        if (line.Contains("Ideal CV"))
                        {
                            idealCV = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Ideal CV"), line.Length - line.LastIndexOf("Ideal CV")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 11:
                        if (line.Contains("Wet CV"))
                        {
                            wetCV = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Dry CV")).Replace("Wet CV", "").Replace("  ", ""));
                        }
                        if (line.Contains("Dry CV"))
                        {
                            dryCV = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Dry CV"), line.Length - line.LastIndexOf("Dry CV")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 12:
                        if (line.Contains("Contract Temp."))
                        {
                            contractTemp = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Contract Press.")).Replace("Contract Temp.", "").Replace("  ", ""));
                        }
                        if (line.Contains("Contract Press."))
                        {
                            contractPress = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Contract Press."), line.Length - line.LastIndexOf("Contract Press.")).Replace("  ", "").Replace("Contract Press.", ""));
                            lineNum++;
                        }
                        break;
                    case 13:
                        if (line.Contains("Number of Cycles"))
                        {
                            numCycles = (int)GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Connected Stream")).Replace("Number of Cycles", "").Replace("  ", ""));
                        }
                        if (line.Contains("Connected Stream"))
                        {
                            connectedStreams = (int)GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Connected Stream"), line.Length - line.LastIndexOf("Connected Stream")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 14:
                        if (line.Contains("Atmospheric Pressure"))
                        {
                            atmoPressure = GetNumbersAndDecimalsAsFloat(line.Replace("Atmospheric Pressure", "").Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    default:
                        break;
                }

            }
            return gasList;
        }

        public bool breaksFileNameRules(string name)
        {
            if ((name.Contains("\\") || name.Contains("/") ||
                    name.Contains(":") || name.Contains("*") ||
                    name.Contains("?") || name.Contains("\"") ||
                    name.Contains("<") || name.Contains(">") ||
                    name.Contains("|"))) return true;
            else return false;
        }

        public string GenerateXMVCSV(string sourceLoc, string outputLoc)
        {
            List<Gas> gaslist = loadData(sourceLoc);
            if (gaslist != null)
            {
                meterDescription = meterDescription.Replace("/", ",").Replace("\\", ",");
                string fileName = outputLoc + "\\" + meterDescription + ".csv";
                if (breaksFileNameRules(fileName.Split('\\')[fileName.Split('\\').Length - 1]))
                {
                    MessageBox.Show("File name breaks Window's rules", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                string fileTXT = idealCV + "\n" + realRelDensity + "\n" +
                    gaslist.Find(x => x.Name == "Carbon-Dioxide").Norm + "\n" + gaslist.Find(x => x.Name == "Nitrogen").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Methane").Norm + "\n" + gaslist.Find(x => x.Name == "Ethane").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Propane").Norm + "\n" + gaslist.Find(x => x.Name == "IsoButane").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Butane").Norm + "\n" +
                    (gaslist.Find(x => x.Name == "IsoPentane").Norm + gaslist.Find(x => x.Name == "NeoPentane").Norm) + "\n" +
                    gaslist.Find(x => x.Name == "Pentane").Norm + "\n" + gaslist.Find(x => x.Name == "Hexanes").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Heptanes").Norm + "\n" + gaslist.Find(x => x.Name == "Octanes").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Nonanes").Norm;
                File.WriteAllText(fileName, fileTXT);
                return fileName;
            }
            else
            {
                MessageBox.Show("File could not be read correctly.", "Format Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public string GenerateRealfloCSV(string sourceLoc, string outputLoc)
        {
            List<Gas> gaslist = loadData(sourceLoc);
            if (gaslist != null)
            {
                meterDescription = meterDescription.Replace("/", ",").Replace("\\",",");
                string fileName = outputLoc + "\\" + meterDescription + ".csv";
                if (breaksFileNameRules(fileName.Split('\\')[fileName.Split('\\').Length - 1]))
                {
                    MessageBox.Show("File name breaks Window's rules", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                string fileTXT =
                    gaslist.Find(x => x.Name == "Methane").Norm + "\n" + gaslist.Find(x => x.Name == "Nitrogen").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Carbon-Dioxide").Norm + "\n" + gaslist.Find(x => x.Name == "Ethane").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Propane").Norm + "\n0\n0\n0\n0\n0\n0\n" + gaslist.Find(x => x.Name == "IsoButane").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Butane").Norm + "\n" +
                    (gaslist.Find(x => x.Name == "IsoPentane").Norm + gaslist.Find(x => x.Name == "NeoPentane").Norm) + "\n" +
                    gaslist.Find(x => x.Name == "Pentane").Norm + "\n" + gaslist.Find(x => x.Name == "Hexanes").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Heptanes").Norm + "\n" + gaslist.Find(x => x.Name == "Octanes").Norm + "\n" +
                    gaslist.Find(x => x.Name == "Nonanes").Norm + "\n0\n0\nl\n" + realRelDensity + "\n" + idealCV;
                File.WriteAllText(fileName, fileTXT);
                return fileName;
            }
            else
            {
                MessageBox.Show("File could not be read correctly.", "Format Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public bool GenerateSpreadsheet1(string sourceLoc, string outputLoc, bool showReport)
        {
            List<Gas> gaslist = loadData(sourceLoc);
            if(gaslist != null)
            {
                    string fileName = outputLoc + "\\" + meterID + "..3.TXT";
                    if (breaksFileNameRules(fileName.Split('\\')[fileName.Split('\\').Length - 1]))
                    {
                        MessageBox.Show("File name breaks Window's rules", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    string dateOnly = printDateTime.Split(' ')[0];
                    string fileTXT = meterID+"\tA\t"+dateOnly+"\t\t"+dateOnly+"\tS"+"\t\t\t"+realRelDensity+"\t\t\t14.7300\t"+
                        gaslist.Find(x => x.Name=="Carbon-Dioxide").Norm+"\t"+gaslist.Find(x => x.Name=="Nitrogen").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Methane").Norm+"\t"+gaslist.Find(x => x.Name=="Ethane").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Propane").Norm+"\t"+gaslist.Find(x => x.Name=="IsoButane").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Butane").Norm+"\t"+gaslist.Find(x => x.Name=="IsoPentane").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Pentane").Norm+"\t"+gaslist.Find(x => x.Name=="NeoPentane").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Hexanes").Norm+"\t"+gaslist.Find(x => x.Name=="Heptanes").Norm+"\t"+
                        gaslist.Find(x => x.Name=="Octanes").Norm+"\t"+gaslist.Find(x => x.Name=="Nonanes").Norm+
                        "\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t"+dryCV+"\t"+gaslist.Find(x => x.Name=="Total").Liquids+"\t\t\t\t"+
                        analyzedBy+"\t\t\t"+flowingPressure+"\t"+flowingTemp+"\t"+gaslist.Find(x => x.Name=="Total").UnNorm+"\t"+
                        compressibility+"\t\t\t\t\t\t\t\t\t\t"+gaslist.Find(x => x.Name=="Ethane").Liquids+"\t"+
                        gaslist.Find(x => x.Name=="Propane").Liquids+"\t"+gaslist.Find(x => x.Name=="IsoButane").Liquids+"\t"+
                        gaslist.Find(x => x.Name=="Butane").Liquids+"\t"+gaslist.Find(x => x.Name=="IsoPentane").Liquids+"\t"+
                        gaslist.Find(x => x.Name=="Pentane").Liquids+"\t"+gaslist.Find(x => x.Name=="Hexanes").Liquids+"\t"+
                        gaslist.Find(x => x.Name=="Heptanes").Liquids+"\t"+wetCV+"\t \t"+dateOnly+"\t\t\tA\t"+atmoPressure+"\t"+
                        gaslist.Find(x => x.Name=="Octanes").Liquids+"\t"+gaslist.Find(x => x.Name=="Nonanes").Liquids+
                        "\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t ";
                    File.WriteAllText(fileName,fileTXT);
                    if(showReport) System.Diagnostics.Process.Start(fileName);
                return true;
            }
            else
            {
                MessageBox.Show("File could not be read correctly.", "Format Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
        }

        public bool GenerateRunReportRename(string sourceLoc, string outputLoc, string meter_id, string meter_desc, bool doAll, bool showReport)
        {
            /*Run report names follow this format:
              meter_id..meter_desc.runNumber.txt
              Spreadshee report names follow this format:
              meter_id..3
            */
            int runs = 1;
            int runNumber;
            string fileName;
            string fileText;
            string[] path = sourceLoc.Split('\\');
            string sourceFileName = path[path.Length - 1];
            string old_id = sourceFileName.Split('.')[0];
            int dirIndex = sourceLoc.LastIndexOf('\\') + 1;
            string dir = sourceLoc.Remove(dirIndex);
            if (Int32.TryParse(sourceFileName.Split('.')[3], out runNumber))
            {
                if (doAll)
                {
                    //set for-loop parameters
                    runs = 3;
                    runNumber = 1;
                    //generate renamed spreadsheet (easier method below)
                    if (File.Exists(dir + old_id + "..3.txt"))
                    {
                        fileText = File.ReadAllText(dir + old_id + "..3.txt");
                        fileText = fileText.Replace(old_id, meter_id);
                        fileName = outputLoc + "\\" + meter_id + "..3.TXT";
                        //Console.WriteLine(fileName.Split('\\')[fileName.Split('\\').Length - 1]);
                        if (breaksFileNameRules(fileName.Split('\\')[fileName.Split('\\').Length - 1]))
                        {
                            MessageBox.Show("Spreadsheet file name breaks Window's rules", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        File.WriteAllText(fileName,fileText);
                        if (showReport) System.Diagnostics.Process.Start(fileName);
                    }
                    else
                    {
                        MessageBox.Show("No such spreadsheet to rename.\nCheck naming scheme.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                //generate run reports
                for (int i = 0; i < runs; i++)
                {
                    sourceLoc = sourceLoc.Remove(sourceLoc.Length - 5) + runNumber.ToString() + ".txt"; //replace run number
                    string newMeterIDLine = "      Meter ID:          " + meter_id;
                    string newMeterDescLine = "                         " + meter_desc + "..." + runNumber;
                    if (File.Exists(sourceLoc))
                    {
                        string[] lines = File.ReadAllLines(sourceLoc);
                        lines[3] = newMeterIDLine;
                        lines[4] = newMeterDescLine;
                        fileName = outputLoc + "\\" + meter_id + ".." + meter_desc.Replace("/", ",").Replace("\\", ",") + "." + runNumber + ".txt";
                        if (breaksFileNameRules(fileName.Split('\\')[fileName.Split('\\').Length - 1]))
                        {
                            MessageBox.Show("Run report file name breaks Window's rules", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        fileText = "";
                        foreach (string line in lines) fileText = fileText + line + "\n";
                        File.WriteAllText(fileName, fileText);
                        if (showReport) System.Diagnostics.Process.Start(fileName);
                        runNumber += 1; 
                    }
                    else
                    {
                        MessageBox.Show("No such report to rename.\nCheck naming scheme.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                //way easier method to generate spreadsheet. won't have as many significant digits though
                //if (doAll)
                //{
                //    bool success = GenerateSpreadsheet1(sourceLoc.Remove(sourceLoc.Length - 5) + "3.txt", outputLoc, showReport);
                //    if (!success)
                //    {
                //        MessageBox.Show("Failed to create spreadsheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //        return false;
                //    } 
                //}
                return true; 
            }
            else
            {
                MessageBox.Show("Check file numbering scheme.", "Error: could not read file(s).", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool GenerateLimerockReport(string sourceLoc,int hexaneCalcType, string outputLoc, bool showReport)
        {
            List<Gas> gasList = loadData(sourceLoc);

            //successfully scraped
            if (gasList != null)
            {
                //Intialize Doc
                PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();
                document.Info.Title = meterID + " Report";
                PdfSharp.Pdf.PdfPage page = document.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Calibri", 11, XFontStyle.Regular);
                XFont bfont = new XFont("Calibri", 11, XFontStyle.Bold);
                XFont lbfont = new XFont("Calibri", 11.5, XFontStyle.Bold);
                XPen greyPen = new XPen(XColors.LightGray, Math.PI);
                
                //Doc Start
                DrawImage(gfx, @"Resources\winn-marion_graphic.PNG", 50,65,190,75);
                
                gfx.DrawString("Sampled By", bfont, XBrushes.Black, new XRect(256, 80, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawString(analyzedBy, bfont, XBrushes.Black, new XRect(310, 80, 85, 20), XStringFormats.Center);

                gfx.DrawString("Date", bfont, XBrushes.Black, new XRect(256, 110, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawString(printDateTime, bfont, XBrushes.Black, new XRect(310, 110, 85, 20), XStringFormats.Center);

                gfx.DrawString("Meter ID", bfont, XBrushes.Black, new XRect(80, 140, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawRectangle(XBrushes.LightGray, 145, 140, 190, 18);
                gfx.DrawString(meterID, lbfont, XBrushes.Black, new XRect(183, 140, 85, 20), XStringFormats.Center);

                gfx.DrawString("Flowing Pressure", bfont, XBrushes.Black, new XRect(60, 165, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawRectangle(XBrushes.LightGray, 145, 165, 60, 18);
                gfx.DrawString(flowingPressure.ToString(), lbfont, XBrushes.Black, new XRect(145, 165, 60, 15), XStringFormats.Center);
                gfx.DrawString("PSIG", font, XBrushes.Black, new XRect(210, 165, 40, 20), XStringFormats.CenterLeft);

                gfx.DrawString("Sample type", bfont, XBrushes.Black, new XRect(250, 165, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawRectangle(XBrushes.LightGray, 310, 165, 40, 18);
                gfx.DrawString(sampleType, lbfont, XBrushes.Black, new XRect(310, 165, 40, 20), XStringFormats.Center);

                gfx.DrawString("Flowing Temp", bfont, XBrushes.Black, new XRect(380, 165, 85, 20), XStringFormats.CenterLeft);
                gfx.DrawRectangle(XBrushes.LightGray, 450, 165, 60, 18);
                gfx.DrawString(flowingTemp.ToString(), lbfont, XBrushes.Black, new XRect(450, 165, 60, 15), XStringFormats.Center);
                gfx.DrawString("Deg F", font, XBrushes.Black, new XRect(510, 165, 85, 20), XStringFormats.CenterLeft);

                gfx.DrawRectangle(XBrushes.LightGray, 20, 220, 85, 400);
                gfx.DrawString("Comp", bfont, XBrushes.Black, new XRect(20, 200, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString("UnNorm %", bfont, XBrushes.Black, new XRect(100, 200, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString("Normal %", bfont, XBrushes.Black, new XRect(180, 200, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString("Liquids GPM", bfont, XBrushes.Black, new XRect(260, 200, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString("Ideal BTU/SCF", bfont, XBrushes.Black, new XRect(340, 200, 85, 20), XStringFormats.CenterRight);

                int yDist = 200;
                int ySteps = 20;
                Gas pentanePlus = new Gas("Pentane+", 0, 0, 0, 0, 0);
                Gas hexanes = new Gas("Hexanes",0,0,0,0,0);
                foreach (Gas substance in gasList)
                {
                    if (substance.Name == "Propane" || substance.Name == "IsoButane" || substance.Name == "IsoPentane" ||
                        substance.Name == "Nitrogen" || substance.Name == "Methane" || substance.Name == "Carbon-Dioxide"
                        || substance.Name == "Ethane" || substance.Name == "Butane")
                    {
                        yDist = yDist + ySteps;
                        gfx.DrawString(substance.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
                    }

                    if (substance.Name == "NeoPentane" || substance.Name == "Pentane")
                    {

                        pentanePlus.UnNorm += substance.UnNorm;
                        pentanePlus.Norm += substance.Norm;
                        pentanePlus.Liquids += substance.Liquids;
                        pentanePlus.Ideal += substance.Ideal;
                        pentanePlus.RelDensity += substance.RelDensity;
                        if (substance.Name == "Pentane")
                        {
                            yDist = yDist + ySteps;
                            gfx.DrawString(pentanePlus.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(pentanePlus.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(pentanePlus.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(pentanePlus.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(pentanePlus.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
                        }
                    }

                    if (hexaneCalcType == 0)

                    {
                        if (substance.Name == "Hexanes" || substance.Name == "Hexane+")
                        {

                            hexanes.UnNorm += substance.UnNorm;
                            hexanes.Norm = substance.Norm;
                            hexanes.Liquids += substance.Liquids;
                            hexanes.Ideal += substance.Ideal;
                            hexanes.RelDensity += substance.RelDensity;
                            if (substance.Name == "Hexane+")//Final Gas!!
                            {
                                yDist = yDist + ySteps;
                                gfx.DrawString(hexanes.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
                            }
                        }
                    }
                    else if (hexaneCalcType == 1)
                    {
                        if (substance.Name == "Hexanes" || substance.Name == "Heptanes" || substance.Name == "Octanes" || substance.Name == "Nonane+" || substance.Name == "Nonanes"
                        || substance.Name == "Decanes" || substance.Name == "Undecanes" /*|| substance.Name == "Pentane-" */|| substance.Name == "Hexane+" /*|| substance.Name == "Propane+" || substance.Name == "Ethane-"*/)

                        {

                            hexanes.UnNorm += substance.UnNorm;
                            hexanes.Norm = substance.Norm;
                            hexanes.Liquids += substance.Liquids;
                            hexanes.Ideal += substance.Ideal;
                            hexanes.RelDensity += substance.RelDensity;
                            if (substance.Name == "Hexane+")//Final Gas!!
                            {
                                yDist = yDist + ySteps;
                                gfx.DrawString(hexanes.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                                gfx.DrawString(hexanes.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
                            }
                        }
                    }
                    

                    if (substance.Name == "Total")
                    {
                        yDist = yDist + ySteps;
                        gfx.DrawRectangle(XBrushes.Blue, 20, yDist, 600, 1);
                        gfx.DrawRectangle(XBrushes.Blue, 20, yDist + ySteps, 600, 1);
                        gfx.DrawString(substance.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                        gfx.DrawString(substance.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
                    }
                }
                gfx.DrawString("Compressibility", bfont, XBrushes.Black, new XRect(20, 440, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(compressibility.ToString(), bfont, XBrushes.Black, new XRect(100, 440, 85, 20), XStringFormats.Center);

                gfx.DrawString("Wet CV", bfont, XBrushes.Black, new XRect(20, 460, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(wetCV.ToString(), bfont, XBrushes.Black, new XRect(100, 460, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", bfont, XBrushes.Black, new XRect(180, 460, 85, 20), XStringFormats.Center);

                gfx.DrawString("Dry CV", bfont, XBrushes.Black, new XRect(20, 480, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(dryCV.ToString(), bfont, XBrushes.Black, new XRect(100, 480, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", bfont, XBrushes.Black, new XRect(180, 480, 85, 20), XStringFormats.Center);

                gfx.DrawString("Real Rel. Density", font, XBrushes.Black, new XRect(320, 440, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(realRelDensity.ToString(), font, XBrushes.Black, new XRect(400, 440, 85, 20), XStringFormats.Center);

                gfx.DrawString("Ideal CV", font, XBrushes.Black, new XRect(320, 460, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(idealCV.ToString(), font, XBrushes.Black, new XRect(400, 460, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", font, XBrushes.Black, new XRect(480, 460, 85, 20), XStringFormats.Center);

                gfx.DrawString("Superior Wobbe", font, XBrushes.Black, new XRect(320, 480, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(superiorWobbe.ToString(), font, XBrushes.Black, new XRect(400, 480, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", font, XBrushes.Black, new XRect(480, 480, 85, 20), XStringFormats.Center);

                //Save Doc
                document.Save(outputLoc + "\\" + meterID + ".pdf");

                //Debug view
                if(showReport) Process.Start(outputLoc + "\\" + meterID + ".pdf");
            }
            else
            {
                MessageBox.Show("Failed to read report data. Check formatting.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            
            return true;
        }
        public class CalData {
            public List<string> FoundTest = new List<string>();
            public List<string> FoundMeter = new List<string>();
            public List<string> LeftTest = new List<string>();
            public List<string> LeftMeter = new List<string>();

        }
        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }
        public string CellValueOrNull(Microsoft.Office.Interop.Excel.Range range,int Y, int X)
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
        public string CellDateOrNull(Microsoft.Office.Interop.Excel.Range range, int Y, int X)
        {
            string dvalue = "";
            string value = CellValueOrNull(range, Y, X);
            if (value != "")
            {
                dvalue = DateTime.FromOADate(double.Parse(value)).ToString();
            }
            return dvalue;
        }
        public bool GenerateExcelCalReport(string sourceLoc, string outputLoc, bool showReport)
        {
            string reportDate, collectionTime, deviceID, system, location, field, state, producer, 
                calibrationTime, purchaser, tagType, spTapLocation, remarks, calibratedBy, calBydate, witness, witnessDate,currentPlateSize, speceficGravity, pipeSize, temperatureBias, amtosPressure, pressureBas, temperatureBase, dpRangeLow, dpRangeHigh, spRangeLow, spRangeHigh, tempRangeLow, tempRangeHigh;
            reportDate = collectionTime = deviceID = system = location = field = state = producer =
                calibrationTime = purchaser = tagType = spTapLocation = remarks = calibratedBy = calBydate = witness = witnessDate = currentPlateSize = speceficGravity = pipeSize = temperatureBias = amtosPressure = pressureBas = temperatureBase = dpRangeLow = dpRangeHigh = spRangeLow = spRangeHigh = tempRangeLow = tempRangeHigh = String.Empty;
            CalData diffData = new CalData();
            CalData diffWSPData = new CalData();
            CalData staticData = new CalData();
            CalData tempData = new CalData();

            //Open Excel and read file---------------------------------------------------------------------------------
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(sourceLoc);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            //Read info here
            //Debug Read all
            /*
            string[,] data = new string[40,60]; //xlRange.Cells[4, 6].Value2.ToString();
            for (int i = 1; i < 60; i++)
            {
                for (int n = 1; n < 40; n++)
                {
                    if (xlRange.Cells[i, n] != null) {
                        if (xlRange.Cells[i, n].Value2 != null)
                        {
                            data[i, n] = xlRange.Cells[i, n].Value2.ToString();
                        }
                    }
                    
                }
            }*/
            reportDate = CellDateOrNull(xlRange,2,6);
            collectionTime = CellDateOrNull(xlRange, 2, 24);
            deviceID = CellValueOrNull(xlRange, 4, 6);
            system = CellValueOrNull(xlRange, 4, 24);
            location = CellValueOrNull(xlRange, 5, 6);
            field = CellValueOrNull(xlRange, 5, 24);
            state = CellValueOrNull(xlRange, 6,24);
            producer = CellValueOrNull(xlRange, 7, 24);
            calibrationTime = CellDateOrNull(xlRange, 8,8);
            purchaser = CellValueOrNull(xlRange, 8,24);
            currentPlateSize = CellValueOrNull(xlRange,10,8);
            tagType = CellValueOrNull(xlRange, 10,19);
            speceficGravity = CellValueOrNull(xlRange, 10,31);
            pipeSize = CellValueOrNull(xlRange, 11,8);
            spTapLocation = CellValueOrNull(xlRange, 11, 19);
            temperatureBias = CellValueOrNull(xlRange, 11, 31);
            amtosPressure = CellValueOrNull(xlRange, 12, 8);
            pressureBas = CellValueOrNull(xlRange, 12, 19);
            temperatureBase = CellValueOrNull(xlRange, 12, 31);

            for (int i = 0; i < 10; i++)
            {
                diffData.FoundTest.Add(CellValueOrNull(xlRange, 21+i, 2));
                diffData.FoundMeter.Add(CellValueOrNull(xlRange, 21 + i, 6));
                diffData.LeftTest.Add(CellValueOrNull(xlRange, 21 + i, 10));
                diffData.LeftMeter.Add(CellValueOrNull(xlRange, 21 + i, 14));

                staticData.FoundTest.Add(CellValueOrNull(xlRange, 34 + i, 2));
                staticData.FoundMeter.Add(CellValueOrNull(xlRange, 34 + i, 6));
                staticData.LeftTest.Add(CellValueOrNull(xlRange, 34 + i, 10));
                staticData.LeftMeter.Add(CellValueOrNull(xlRange, 34 + i, 14));

                tempData.FoundTest.Add(CellValueOrNull(xlRange, 34 + i, 18));
                tempData.FoundMeter.Add(CellValueOrNull(xlRange, 34 + i, 22));
                tempData.LeftTest.Add(CellValueOrNull(xlRange, 34 + i, 26));
                tempData.LeftMeter.Add(CellValueOrNull(xlRange, 34 + i, 30));
            }

            for (int i = 0; i < 5; i++)
            {
                diffWSPData.FoundTest.Add(CellValueOrNull(xlRange, 21 + i, 18));
                diffWSPData.FoundMeter.Add(CellValueOrNull(xlRange, 21 + i, 22));
                diffWSPData.LeftTest.Add(CellValueOrNull(xlRange, 21 + i, 26));
                diffWSPData.LeftMeter.Add(CellValueOrNull(xlRange, 21 + i, 30));
            }

            dpRangeLow = CellValueOrNull(xlRange, 44,2 );
            dpRangeHigh = CellValueOrNull(xlRange, 45, 2);
            spRangeLow = CellValueOrNull(xlRange, 44, 6);
            spRangeHigh = CellValueOrNull(xlRange, 45, 6);
            tempRangeLow = CellValueOrNull(xlRange, 44, 10);
            tempRangeHigh = CellValueOrNull(xlRange, 45, 10);

            remarks = CellValueOrNull(xlRange, 47, 2);

            calibratedBy = CellValueOrNull(xlRange, 55, 7);
            calBydate = CellDateOrNull(xlRange, 55, 30);
            witness = CellValueOrNull(xlRange, 57, 7);
            witnessDate = CellDateOrNull(xlRange, 57, 30);

            //close Excel
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Release Excel
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            //close
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlApp);

            //quit
            //xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //-------------------------------------------------------------------------------------------------------------

            iText.Kernel.Pdf.PdfDocument pdfDoc =
                new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(System.AppDomain.CurrentDomain.BaseDirectory + "Resources\\Calibration Doc 250 1000 PSI.pdf"),
                new iText.Kernel.Pdf.PdfWriter(outputLoc + "\\" + Path.GetFileName(sourceLoc).Split('.')[0] + "_CalibrationReport.pdf"));// "C:\\Users\\193039\\Desktop\\Calibration Doc 250 1000 PSI_Test.pdf"));
            PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);
            form.SetGenerateAppearance(true);
            PdfFont font = PdfFontFactory.CreateFont(FONT, PdfEncodings.IDENTITY_H);

            
            IDictionary<string, iText.Forms.Fields.PdfFormField> fields = form.GetFormFields();
            string[] pdfKeys = { "Other","ORIFICE METER","ELECTRONIC FLOW METER","OTHER","AVG DP","Diff OK","MeterRow1","MeterRow1_2","LeftDPURow1",
                "DeviceRow1_5","MeterRtdRow1","DeviceRow1_6","MeterRtdRow1_2","DeviceRow2_5","MeterRtdRow2","DeviceRow2_6","MeterRtdRow2_2","DeviceRow3_5",
                "MeterRtdRow3","DeviceRow3_6","MeterRtdRow3_2","IDTube Size","Plate out","Plate Size","AS FOUND FLOW","AS LEFT FLOW","Witness By","Text1",
                "Check Box2","Check Box3","Text5","Text6","Text7","Text8","Text9","Text10","Text12","Check Box13","Check Box14","Check Box15","Check Box16",
                "Check Box17","Check Box18","Text19","Text20","Text21","Text22","Text23","Text24","Text25","Text26","Check Box4","Check Box5","Check Box6",
                "Check Box7","Check Box8","Check Box9","Check Box10","Check Box11","Check Box12","Check Box19","Check Box20","Check Box21","Text27","Text28",
                "Text29","Check Box30","Check Box31","Text2","Text30","Check Box32","Check Box33","Text3","Text4","Check Box22","Check Box23","Text31",
                "Check Box34","Check Box35","Text13","Check Box24","Check Box25","Text32","Text33","Text34","DeviceRow1_2","DeviceRow1_2.0","DeviceRow1_2.0.1",
                "DeviceRow1_2.0.0","DeviceRow1_2.0.0.0","DeviceRow1_2.0.0.0.0","DeviceRow1_2.0.0.0.1","DeviceRow1_2.0.0.1","DeviceRow1_2.0.0.1.0",
                "DeviceRow1_2.0.0.1.1","DeviceRow1_2.0.0.2","DeviceRow1_2.0.0.2.0","DeviceRow1_2.0.0.2.1","DeviceRow1_2.0.0.3","DeviceRow1_2.0.0.3.0",
                "DeviceRow1_2.0.0.3.1","DeviceRow1_2.0.0.4","DeviceRow1_2.0.0.4.0","DeviceRow1_2.0.0.4.1","DeviceRow1_2.0.0.5","DeviceRow1_2.0.0.5.0",
                "DeviceRow1_2.0.0.5.1","DeviceRow1_2.0.0.6","DeviceRow1_2.0.0.6.0","DeviceRow1_2.0.0.6.1","DeviceRow1","DeviceRow1.0","DeviceRow1.0.0",
                "DeviceRow1.0.1","DeviceRow1.1","DeviceRow1.1.0","DeviceRow1.1.1","DeviceRow1.2","DeviceRow1.2.0","DeviceRow1.2.1","DeviceRow1.3",
                "DeviceRow1.3.0","DeviceRow1.3.1","DeviceRow1.4","DeviceRow1.4.0","DeviceRow1.4.1","DeviceRow1.5","DeviceRow1.5.0","DeviceRow1.5.1",
                "DeviceRow1.6","DeviceRow1.6.0","DeviceRow1.6.1","MeterRow1_3","MeterRow1_3.0","MeterRow1_3.0.0","MeterRow1_3.0.2","MeterRow1_3.0.2.0",
                "MeterRow1_3.0.2.1","MeterRow1_3.0.2.2","MeterRow1_3.0.2.3","MeterRow1_3.0.2.4","MeterRow1_3.0.2.5","MeterRow1_3.0.2.6","MeterRow1_3.0.2.7",
                "MeterRow1_3.1","MeterRow1_3.1.0","MeterRow1_3.1.1","MeterRow1_3.1.1.0","MeterRow1_3.1.1.3","MeterRow1_3.1.1.4","MeterRow1_3.1.1.5",
                "MeterRow1_3.1.1.6","MeterRow1_3.1.1.7","MeterRow1_3.1.1.1","MeterRow1_3.1.1.1.0","MeterRow1_3.1.1.1.1","MeterRow1_3.2","MeterRow1_3.2.0",
                "MeterRow1_3.3","MeterRow1_3.3.0","MeterRow1_3.4","MeterRow1_3.4.0","MeterRow1_3.5","MeterRow1_3.5.0","MeterRow1_3.6","MeterRow1_3.6.0",
                "MeterRow1_3.7","MeterRow1_3.7.0","Differential UpRow2","Differential UpRow2.0","Differential UpRow2.1","Differential UpRow2.2",
                "Differential UpRow2.3","Differential UpRow2.4","Differential UpRow2.5","Differential UpRow2.6","Differential UpRow2.7",
};
            float fontSize = 12f;
            form.GetField("Text10").SetValue(deviceID, font, fontSize); //Station 35
            form.GetField("Text9").SetValue(location, font, fontSize); //Location 34
            form.GetField("Text30").SetValue(reportDate, font, fontSize); //Date 69
            form.GetField("Text1").SetValue(producer, font, fontSize); //Company 27
            form.GetField("Text7").SetValue(field, font, fontSize); //Field 32
            form.GetField("Text8").SetValue(state, font, fontSize);//State 33
            //form.GetField("Text3").SetValue(, font, fontSize); //Lease ID 72
            //form.GetField("Text4").SetValue(deviceID, font, fontSize);//API 73
            //form.GetField("Other").SetValue(deviceID, font, fontSize);//Cal Freq Other 1
            //form.GetField("ORIFICE METER").SetValue(, font, fontSize); //ORIFICE METER 1
            form.GetField("ELECTRONIC FLOW METER").SetValue("Yes", font, fontSize);//EFM 2
            //form.GetField("Text2").SetValue(, font, fontSize);//Last Verif Date 68
            form.GetField("Text29").SetValue("ABB", font, fontSize);//meter make 65
            form.GetField("Text26").SetValue("XMV", font, fontSize);//model 50
            form.GetField("Text25").SetValue(deviceID, font, fontSize);//SN 49
            //form.GetField("OTHER").SetValue(, font, fontSize);//EFM Type Other 3
            //form.GetField("Text23").SetValue(, font, fontSize);//Diff Range 47
            //form.GetField("Text24").SetValue(, font, fontSize);//Stat Range 48
            //form.GetField("Text22").SetValue(, font, fontSize);//Diff Found 46
            //form.GetField("Text20").SetValue(, font, fontSize);//Diff Left 44
            //form.GetField("Text21").SetValue(, font, fontSize);//gas gravity 45
            //form.GetField("AVG DP").SetValue(, font, fontSize);//Avg DP 4
            //form.GetField("Diff OK").SetValue(, font, fontSize);//Diff Okay 5
            form.GetField("Text10").SetValue(deviceID, font, fontSize);
            //tables
            form.GetField("MeterRow1_3.1.1.0").SetValue(diffData.FoundTest[0], font, fontSize);//147
            form.GetField("MeterRow1_3.1.1.1.0").SetValue(diffData.FoundTest[1], font, fontSize);//154
            form.GetField("MeterRow1_3.1.1.1.1").SetValue(diffData.FoundTest[2], font, fontSize);//155
            form.GetField("MeterRow1_3.1.1.3").SetValue(diffData.FoundTest[3], font, fontSize);//148
            form.GetField("MeterRow1_3.1.1.4").SetValue(diffData.FoundTest[4], font, fontSize);//149
            form.GetField("MeterRow1_3.1.1.5").SetValue(diffData.FoundTest[5], font, fontSize);//150
            form.GetField("MeterRow1_3.1.1.6").SetValue(diffData.FoundTest[6], font, fontSize);//151
            form.GetField("MeterRow1_3.1.1.7").SetValue(diffData.FoundTest[7], font, fontSize);//152

            form.GetField("MeterRow1_3.0.0").SetValue(diffData.FoundMeter[0], font, fontSize);//134
            form.GetField("MeterRow1_3.1.0").SetValue(diffData.FoundMeter[1], font, fontSize);//145
            form.GetField("MeterRow1_3.2.0").SetValue(diffData.FoundMeter[2], font, fontSize);//157
            form.GetField("MeterRow1_3.3.0").SetValue(diffData.FoundMeter[3], font, fontSize);//159
            form.GetField("MeterRow1_3.4.0").SetValue(diffData.FoundMeter[4], font, fontSize);//161
            form.GetField("MeterRow1_3.5.0").SetValue(diffData.FoundMeter[5], font, fontSize);//163
            form.GetField("MeterRow1_3.6.0").SetValue(diffData.FoundMeter[6], font, fontSize);//165
            form.GetField("MeterRow1_3.7.0").SetValue(diffData.FoundMeter[7], font, fontSize);//167

            form.GetField("Differential UpRow2.0").SetValue(diffData.LeftTest[0], font, fontSize);//169
            form.GetField("Differential UpRow2.1").SetValue(diffData.LeftTest[1], font, fontSize);//170
            form.GetField("Differential UpRow2.2").SetValue(diffData.LeftTest[2], font, fontSize);//171
            form.GetField("Differential UpRow2.3").SetValue(diffData.LeftTest[3], font, fontSize);//172
            form.GetField("Differential UpRow2.4").SetValue(diffData.LeftTest[4], font, fontSize);//173
            form.GetField("Differential UpRow2.5").SetValue(diffData.LeftTest[5], font, fontSize);//174
            form.GetField("Differential UpRow2.6").SetValue(diffData.LeftTest[6], font, fontSize);//175
            form.GetField("Differential UpRow2.7").SetValue(diffData.LeftTest[7], font, fontSize);//176

            form.GetField("MeterRow1_3.0.2.0").SetValue(diffData.LeftMeter[0], font, fontSize);//136
            form.GetField("MeterRow1_3.0.2.1").SetValue(diffData.LeftMeter[1], font, fontSize);//137
            form.GetField("MeterRow1_3.0.2.2").SetValue(diffData.LeftMeter[2], font, fontSize);//138
            form.GetField("MeterRow1_3.0.2.3").SetValue(diffData.LeftMeter[3], font, fontSize);//139
            form.GetField("MeterRow1_3.0.2.4").SetValue(diffData.LeftMeter[4], font, fontSize);//140
            form.GetField("MeterRow1_3.0.2.5").SetValue(diffData.LeftMeter[5], font, fontSize);//141
            form.GetField("MeterRow1_3.0.2.6").SetValue(diffData.LeftMeter[6], font, fontSize);//142
            form.GetField("MeterRow1_3.0.2.7").SetValue(diffData.LeftMeter[7], font, fontSize);//143

            form.GetField("Text6").SetValue(amtosPressure, font, fontSize);//Abs Atmo Pressure 31
            //form.GetField("Text5").SetValue(, font, fontSize);//Avg Flow Pressure 30

            form.GetField("DeviceRow1.0.0").SetValue(staticData.FoundTest[0], font, fontSize);//112
            form.GetField("DeviceRow1.1.0").SetValue(staticData.FoundTest[1], font, fontSize);//115
            form.GetField("DeviceRow1.2.0").SetValue(staticData.FoundTest[2], font, fontSize);//118
            form.GetField("DeviceRow1.3.0").SetValue(staticData.FoundTest[3], font, fontSize);//121
            form.GetField("DeviceRow1.4.0").SetValue(staticData.FoundTest[4], font, fontSize);//124
            form.GetField("DeviceRow1.5.0").SetValue(staticData.FoundTest[5], font, fontSize);//127
            form.GetField("DeviceRow1.6.0").SetValue(staticData.FoundTest[6], font, fontSize);//130

            form.GetField("DeviceRow1.0.1").SetValue(staticData.FoundMeter[0], font, fontSize);//113
            form.GetField("DeviceRow1.1.1").SetValue(staticData.FoundMeter[1], font, fontSize);//116
            form.GetField("DeviceRow1.2.1").SetValue(staticData.FoundMeter[2], font, fontSize);//119
            form.GetField("DeviceRow1.3.1").SetValue(staticData.FoundMeter[3], font, fontSize);//122
            form.GetField("DeviceRow1.4.1").SetValue(staticData.FoundMeter[4], font, fontSize);//125
            form.GetField("DeviceRow1.5.1").SetValue(staticData.FoundMeter[5], font, fontSize);//128
            form.GetField("DeviceRow1.6.1").SetValue(staticData.FoundMeter[6], font, fontSize);//131

            form.GetField("DeviceRow1_2.0.0.0.0").SetValue(staticData.LeftTest[0], font, fontSize);//90
            form.GetField("DeviceRow1_2.0.0.1.0").SetValue(staticData.LeftTest[1], font, fontSize);//93
            form.GetField("DeviceRow1_2.0.0.2.0").SetValue(staticData.LeftTest[2], font, fontSize);//96
            form.GetField("DeviceRow1_2.0.0.3.0").SetValue(staticData.LeftTest[3], font, fontSize);//99
            form.GetField("DeviceRow1_2.0.0.4.0").SetValue(staticData.LeftTest[4], font, fontSize);//102
            form.GetField("DeviceRow1_2.0.0.5.0").SetValue(staticData.LeftTest[5], font, fontSize);//105
            form.GetField("DeviceRow1_2.0.0.6.0").SetValue(staticData.LeftTest[6], font, fontSize);//108

            form.GetField("DeviceRow1_2.0.0.0.1").SetValue(staticData.LeftMeter[0], font, fontSize);//91
            form.GetField("DeviceRow1_2.0.0.1.1").SetValue(staticData.LeftMeter[1], font, fontSize);//94
            form.GetField("DeviceRow1_2.0.0.2.1").SetValue(staticData.LeftMeter[2], font, fontSize);//97
            form.GetField("DeviceRow1_2.0.0.3.1").SetValue(staticData.LeftMeter[3], font, fontSize);//100
            form.GetField("DeviceRow1_2.0.0.4.1").SetValue(staticData.LeftMeter[4], font, fontSize);//103
            form.GetField("DeviceRow1_2.0.0.5.1").SetValue(staticData.LeftMeter[5], font, fontSize);//106
            form.GetField("DeviceRow1_2.0.0.6.1").SetValue(staticData.LeftMeter[6], font, fontSize);//109

            //form.GetField("Text31").SetValue(, font, fontSize);//1cur temp 76
            //form.GetField("Text28").SetValue(, font, fontSize);//tmp range 64
            form.GetField("Text12").SetValue(temperatureBase, font, fontSize);//base temp 36

            form.GetField("DeviceRow1_5").SetValue(tempData.FoundTest[0], font, fontSize);//9
            form.GetField("DeviceRow2_5").SetValue(tempData.FoundTest[1], font, fontSize);//13
            form.GetField("DeviceRow3_5").SetValue(tempData.FoundTest[2], font, fontSize);//17

            form.GetField("MeterRtdRow1").SetValue(tempData.FoundMeter[0], font, fontSize);//10
            form.GetField("MeterRtdRow2").SetValue(tempData.FoundMeter[1], font, fontSize);//14
            form.GetField("MeterRtdRow3").SetValue(tempData.FoundMeter[2], font, fontSize);//18

            form.GetField("DeviceRow1_6").SetValue(tempData.LeftTest[0], font, fontSize);//11
            form.GetField("DeviceRow2_6").SetValue(tempData.LeftTest[1], font, fontSize);//15
            form.GetField("DeviceRow3_6").SetValue(tempData.LeftTest[2], font, fontSize);//19

            form.GetField("MeterRtdRow1_2").SetValue(tempData.LeftMeter[0], font, fontSize);//12
            form.GetField("MeterRtdRow2_2").SetValue(tempData.LeftMeter[1], font, fontSize);//16
            form.GetField("MeterRtdRow3_2").SetValue(tempData.LeftMeter[2], font, fontSize);//20

            //form.GetField("Text33").SetValue(, font, fontSize);//Line 1 Notes 83
            //form.GetField("Text34").SetValue(, font, fontSize);//Line 2 Notes 84
            form.GetField("IDTube Size").SetValue(pipeSize, font, fontSize);//ID Tube Size 21
            //form.GetField("Plate out").SetValue(, font, fontSize);//Plate out 22 ?
            form.GetField("Plate Size").SetValue(currentPlateSize, font, fontSize);//Plate Size 23
            //form.GetField("Text32").SetValue(, font, fontSize);//other 82 
            string remarks1, remarks2;
            remarks1 = remarks2 = "";
            if (remarks.Count() > 90)
            {
                remarks1 = remarks.Substring(0, 90);
                remarks2 = remarks.Substring(90, remarks.Count());
            }
            else {
                remarks1 = remarks;
            }
            form.GetField("Text19").SetValue(remarks1, font, fontSize);//Line 1 Remarks 43
            form.GetField("Text13").SetValue(remarks2, font, fontSize);//Line 2 Remarks 79
            //form.GetField("AS FOUND FLOW").SetValue(, font, fontSize);//As FOund Flow 24
            //form.GetField("AS LEFT FLOW").SetValue(, font, fontSize);//As Left FLow 25
            form.GetField("Witness By").SetValue(witness, font, fontSize);//Witness By 26
            form.GetField("Text27").SetValue(calibratedBy, font, fontSize);//inspected by 63


            /*
             * //Debug stuff
            int num = 0;
            foreach (string text in pdfKeys)
            {
                
                form.GetField(pdfKeys[num]).SetValue(num.ToString(), font, 12f);
                num++;
            }
           // form.GetField("Other").SetValue(1,font,12f);
            //int stop = 0;
            */

            pdfDoc.Close();

            //Debug view
            if(showReport) Process.Start(outputLoc + "\\" + Path.GetFileName(sourceLoc).Split('.')[0] + "_CalibrationReport.pdf");
            return true;
        }
    }
}
