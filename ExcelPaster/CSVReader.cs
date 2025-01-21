using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class CSVReader
    {
        private List<List<String>> ArrayStorage = new List<List<string>>();
        private List<String> CurLineList; 
        private String CurCell = "";


        public List<List<String>> GetArrayStorage()
        {
            return ArrayStorage;
        }
        public void ParseCSV(string csv,string customEndline)
        {
            using (var fs = new FileStream(csv, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var sr = new StreamReader(fs);
               ParseCSV(sr,customEndline);
            }
        }
        private void ReadCharacter(char c)
        {
            if (c != ',')
            {
                CurCell += c;
            }
            else//save cell and wipe it clean for next
            {
                CurLineList.Add(CurCell);
                CurCell = "";
            }
        }
        private void ParseCSV(StreamReader stream,string customEndline)
        {
            string curline = "";
            while (curline != null)
            {
                curline = stream.ReadLine();
                if (curline != null)
                {
                    if (customEndline != "")
                    {
                        
                        while (!curline.Contains(customEndline))
                        {
                            curline += stream.ReadLine();
                        }
                    }
                }
              

                if (curline != null)//end of file?
                {
                    CurLineList = new List<string>();
                    foreach (char c in curline)
                    {
                        ReadCharacter(c);
                    }
                    CurLineList.Add(CurCell);
                    CurCell = "";
                    ArrayStorage.Add(CurLineList);
                }
                else
                {
                    break;
                }
            }
           

            
        }
    }
}
