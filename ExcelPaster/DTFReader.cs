using Org.BouncyCastle.Operators;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class DTFReader
    {
        public List<DTFPoint> dtfPoints = new List<DTFPoint>();
        public string sourceFileName = "";
        public void ExtractRegisters(string sourceFileLoc)
        {
           FileInfo f = new FileInfo(sourceFileLoc);
            sourceFileName = f.Name;

            using (var fs = new FileStream(sourceFileLoc, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var sr = new StreamReader(fs);
                string curline = "";

                string curGroupElement = "";
                string curGroupNiceName = "";
                string curPointElement = "";
                string curPointDesc = "";
                string curPointRegNum = "";
                string curPointType = "";
                string curPointReadOnly = "";
                string curPointUDC = "";
                int lineCount = 1;
                bool inComment = false;
                while (curline != null)
                {
                    curline = sr.ReadLine();
                    string[] curlinewords = curline.Split(' ');
                    if (curline != null)
                    {
                        //Comment detected
                        if (curline.Contains("<!--"))
                            inComment = true;
                        if (curline.Contains("-->"))
                            inComment = false;
                        if (inComment)
                        {
                            lineCount++;
                            continue;
                        }
                            

                        //Group Detected
                        if (curline.Contains("niceName"))
                        {
                            curGroupElement = curlinewords[0].Replace("<", "");
                            curGroupNiceName = ReadPropValue(curline, "niceName");
                        }

                        //Point Detected
                        if (curline.Contains("regNum"))
                        {

                               
                            curPointElement = curlinewords[0].Replace("<", "");
                            curPointDesc = ReadPropValue(curline, "desc");
                            curPointRegNum = ReadPropValue(curline, "regNum");
                            curPointType = ReadPropValue(curline, "type");
                            curPointReadOnly = ReadPropValue(curline, "readOnly");
                            curPointUDC = ReadPropValue(curline, "udc");

                            if (lineCount == 1209)
                            {
                                int stop = 1;

                            }
                            dtfPoints.Add(new DTFPoint(curGroupElement, curGroupNiceName, curPointElement, curPointDesc, curPointRegNum, curPointType, curPointReadOnly, curPointUDC, lineCount));
                        }
                        //End of doc detected
                        if (curline.Contains("</deviceDefinition>"))
                            break;
                        lineCount++;
                    }
                }
            }
        }

        private string ReadPropValue(string line, string property)
        {
            if (line.Contains(property))
            {
                int indexNiceName = line.IndexOf(property) + property.Length;
                int indexEquals = line.IndexOf("=", indexNiceName) + 1;
                int indexQuote = line.IndexOf('"', indexEquals) + 1;
                return line.Substring(indexQuote).Split('"')[0];
            }
            else 
            {
                return "";
            }
            
        }

        public void SaveRegisters(string outputFolderLoc, bool displayOuput)
        {
            if (dtfPoints.Count != 0)
            {
                string fileName = outputFolderLoc + "\\" + sourceFileName + ".csv";
                string fileTXT = "Line, GroupElement, GroupNiceName, PointElement, PointDesc, PointRegNum, PointType, PointReadOnly,PointUDC \n";
                foreach (DTFPoint point in dtfPoints)
                {
                    fileTXT += point.lineCount.ToString() + ", " + point.curGroupElement + ", " + point.curGroupNiceName + ", " + point.curPointElement +
                        ", " + point.curPointDesc + ", " + point.curPointRegNum + ", " + point.curPointType + ", " + point.curPointReadOnly + ", " + point.curPointUDC + "\n";

                        
                }
                File.WriteAllText(fileName, fileTXT);

                if (displayOuput) Process.Start(fileName);
                
            }
        }
    }
    public class DTFPoint
    {
        public string curGroupElement;
        public string curGroupNiceName;
        public string curPointElement;
        public string curPointDesc;
        public string curPointRegNum;
        public string curPointType;
        public string curPointReadOnly;
        public string curPointUDC;
        public int lineCount;
        public DTFPoint(string curGroupElement, string curGroupNiceName, string curPointElement, string curPointDesc, string curPointRegNum, string curPointType, string curPointReadOnly, string curPointUDC, int lineCount)
        {
            this.curGroupElement = curGroupElement;
            this.curGroupNiceName = curGroupNiceName;
            this.curPointElement = curPointElement;
            this.curPointDesc = curPointDesc;
            this.curPointRegNum = curPointRegNum;
            this.curPointType = curPointType;
            this.curPointReadOnly = curPointReadOnly;
            this.curPointUDC = curPointUDC;
            this.lineCount = lineCount;
        }
    }

   
}
