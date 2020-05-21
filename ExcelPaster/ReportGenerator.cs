using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Diagnostics;

namespace ExcelPaster
{
    class ReportGenerator
    {

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
        public class Gas{
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
        public bool GenerateLimerockReport(string sourceLoc, string outputLoc)
        {
            string printDateTime="", analyzedBy="", meterID="", analysisTime="", sampleType="", elevation = "";
            float flowingTemp=0, flowingPressure=0, calibrationElevation=0,
                locationElevation=0,inferiorWobbe=0,superiorWobbe=0, 
                compressibility=0, density=0, realRelDensity=0, idealCV=0,wetCV=0,dryCV=0,contractTemp=0,contractPress=0,atmoPressure=0;
            int numCycles=0, connectedStreams=0;
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
                        if (line.Contains("Print Date Time:")) { printDateTime = line.Replace("Print Date Time:", "").Replace("  ", "");
                            lineNum++; }
                        break;
                    case 1:
                        if (line.Contains("Analyzed By:")) { analyzedBy = line.Replace("Analyzed By:", "").Replace("  ", "");
                            lineNum++; }
                        break;
                    case 2:
                        if (line.Contains("Meter ID:")) { meterID = line.Replace("Meter ID:", "").Replace("  ", "").TrimEnd('.');
                            lineNum++; }
                        break;
                    case 3:
                        if (line.Contains("Analysis Time:")) {
                            analysisTime = line.Substring(0, line.LastIndexOf("Sample Type:")).Replace("Analysis Time:", "").Replace("  ", "");
                        }
                        if (line.Contains("Sample Type:")) {
                            sampleType = line.Substring(line.LastIndexOf("Sample Type:"),line.Length - line.LastIndexOf("Sample Type:")).Replace("  ", "").Replace("Sample Type:", "");
                            lineNum++;
                        }
                        break;
                    case 4:
                        if (line.Contains("Flowing Temp.:")) {
                            flowingTemp = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Flowing Pressure:")).Replace("Flowing Temp.:", "").Replace("  ", "").Replace("Deg. F", ""));
                        }
                        if (line.Contains("Flowing Pressure:"))
                        {
                            flowingPressure = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Flowing Pressure:"), line.Length - line.LastIndexOf("Flowing Pressure:")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 5:
                        if (line.Contains("Calibration Elevation:"))
                        {
                            calibrationElevation = GetNumbersAndDecimalsAsFloat(line.Substring(0, line.LastIndexOf("Location Elevation:")).Replace("Calibration Elevation:", "").Replace("  ", ""));
                        }
                        if (line.Contains("Location Elevation:"))
                        {
                            locationElevation = GetNumbersAndDecimalsAsFloat(line.Substring(line.LastIndexOf("Location Elevation:"), line.Length - line.LastIndexOf("Location Elevation:")).Replace("  ", ""));
                            lineNum++;
                        }
                        break;
                    case 6://into table
                        if (tablePrimer == false) {
                            if (line.Contains("----------------------------------------------------------------------------")) { tablePrimer = true; }
                        }else{
                            string combineWords = line.Replace("Carbon Dioxide", "Carbon-Dioxide");
                            string reduceSpaces = string.Join(" ", combineWords.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
                            string[] data = reduceSpaces.Split(' ');
                            if (data.Length >= 5) {
                                gasList.Add(new Gas(data[0], float.Parse(data[1]), float.Parse(data[2]),
                                    float.Parse(data[3]), float.Parse(data[4]), float.Parse(data[5])));
                            }
                            if (data[0] == "Total") { lineNum++; }
                        }
                        break;
                    case 7:
                        if (line.Contains("Elevation"))
                        {
                            elevation = line.Replace("Elevation", "").Replace("  ", "");
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
            //successfully scraped
            if (lineNum == 15)
            {
                //Intialize Doc
                PdfDocument document = new PdfDocument();
                document.Info.Title = meterID + " Report";
                PdfPage page = document.AddPage();
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
                        || substance.Name == "Ethane")
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

                    if (substance.Name == "Hexanes" || substance.Name == "Heptanes" || substance.Name == "Octanes" || substance.Name == "Nonane+" || substance.Name == "Nonanes"
                        || substance.Name == "Decanes" || substance.Name == "Undecanes" || substance.Name == "Pentane-" || substance.Name == "Hexane+" || substance.Name == "Propane+" || substance.Name == "Ethane-")
                    {

                        hexanes.UnNorm += substance.UnNorm;
                        hexanes.Norm += substance.Norm;
                        hexanes.Liquids += substance.Liquids;
                        hexanes.Ideal += substance.Ideal;
                        hexanes.RelDensity += substance.RelDensity;
                        if (substance.Name == "Ethane-")
                        {
                            yDist = yDist + ySteps;
                            gfx.DrawString(hexanes.Name, bfont, XBrushes.Black, new XRect(20, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(hexanes.UnNorm.ToString(), font, XBrushes.Black, new XRect(100, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(hexanes.Norm.ToString(), bfont, XBrushes.Black, new XRect(180, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(hexanes.Liquids.ToString(), font, XBrushes.Black, new XRect(260, yDist, 85, 20), XStringFormats.CenterRight);
                            gfx.DrawString(hexanes.Ideal.ToString(), bfont, XBrushes.Black, new XRect(340, yDist, 85, 20), XStringFormats.CenterRight);
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

                gfx.DrawString("Real Rel. Density", font, XBrushes.Black, new XRect(20, 460, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(realRelDensity.ToString(), font, XBrushes.Black, new XRect(100, 460, 85, 20), XStringFormats.Center);

                gfx.DrawString("Wet CV", bfont, XBrushes.Black, new XRect(20, 480, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(wetCV.ToString(), bfont, XBrushes.Black, new XRect(100, 480, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", bfont, XBrushes.Black, new XRect(180, 480, 85, 20), XStringFormats.Center);

                gfx.DrawString("Ideal CV", font, XBrushes.Black, new XRect(20, 500, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(idealCV.ToString(), font, XBrushes.Black, new XRect(100, 500, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", font, XBrushes.Black, new XRect(180, 500, 85, 20), XStringFormats.Center);

                gfx.DrawString("Dry CV", bfont, XBrushes.Black, new XRect(20, 520, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(dryCV.ToString(), bfont, XBrushes.Black, new XRect(100, 520, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", bfont, XBrushes.Black, new XRect(180, 520, 85, 20), XStringFormats.Center);

                gfx.DrawString("Superior Wobbe", font, XBrushes.Black, new XRect(20, 540, 85, 20), XStringFormats.CenterRight);
                gfx.DrawString(superiorWobbe.ToString(), font, XBrushes.Black, new XRect(100, 540, 85, 20), XStringFormats.Center);
                gfx.DrawString("Btu/SCF", font, XBrushes.Black, new XRect(180, 540, 85, 20), XStringFormats.Center);

                //Save Doc
                document.Save(outputLoc + "/" + meterID + ".pdf");

                //Debug view
                Process.Start(outputLoc + "/" + meterID + ".pdf");
            }
            
            return true;
        }
        
    }
}
