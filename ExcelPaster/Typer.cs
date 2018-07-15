using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsInput;

namespace ExcelPaster
{
    public class Typer
    {
        //private Keyboard kb = new Keyboard();
        private int strokeDelay = 500;
        private InputSimulator In_sim = new InputSimulator();
        public InputHelper ih = new InputHelper();
        public void TypeCSVtoText(List<List<String>> csv, System.ComponentModel.BackgroundWorker bg)
        {
            // ih.LoadDriver();
            for (int i = 0; i < csv.Count(); i++)
            {
                List<string> line = csv[i];
                for (int j = 0; j < line.Count(); j++)
                {
                    string cell = line[j];
                    for (int k = 0; k < cell.Count(); k++)
                    {
                        if (bg.CancellationPending)
                        {
                            break;
                        }
                        char c = cell[k];
                        //SendKey(c);
                        ih.SendKey(c);
                    }
                    if (bg.CancellationPending)
                    {
                        break;
                    }
                    if (j < line.Count() - 1)
                    {
                        // NewCell();
                        ih.SendKey(Interceptor.Keys.Tab);
                    }
                }
                if (bg.CancellationPending)
                {
                    break;
                }
                if (i < csv.Count())
                {
                    //NewLine();
                    ih.SendKey(Interceptor.Keys.Enter);
                }
            }

            ih.UnloadDriver();
        }
        public void TypeCSVtoExcel(List<List<String>> csv, System.ComponentModel.BackgroundWorker bg)
        {
            // ih.LoadDriver();
            for (int i = 0; i < csv.Count(); i++)
            {
                List<string> line = csv[i];
                for (int j = 0; j < line.Count(); j++)
                {
                    string cell = line[j];
                    for (int k = 0; k < cell.Count(); k++)
                    {
                        if (bg.CancellationPending)
                        {
                            break;
                        }
                        char c = cell[k];
                        //SendKey(c);
                        ih.SendKey(c);
                    }
                    if (bg.CancellationPending)
                    {
                        break;
                    }
                    if (j < line.Count() - 1)
                    {
                        // NewCell();
                        ih.SendKey(Interceptor.Keys.Tab);
                    }
                }
                if (bg.CancellationPending)
                {
                    break;
                }
                if (i < csv.Count())
                {
                    //NewLine();
                    ih.SendKey(Interceptor.Keys.Enter);
                }
            }

            ih.UnloadDriver();
        }
        public void TypeCSVtoPCCU(List<List<String>> csv, System.ComponentModel.BackgroundWorker bg)
        {
            // ih.LoadDriver();
            for (int i = 0; i < csv.Count(); i++)
            {
                List<string> line = csv[i];
                for (int j = 0; j < line.Count(); j++)
                {
                    string cell = line[j];
                    for (int k = 0; k < cell.Count(); k++)
                    {
                        if (bg.CancellationPending)
                        {
                            break;
                        }
                        char c = cell[k];
                        //SendKey(c);
                        ih.SendKey(c);
                    }
                    if (bg.CancellationPending)
                    {
                        break;
                    }
                    if (j < line.Count() - 1)
                    {
                        // NewCell();
                        //ih.SendKey(Interceptor.Keys.Tab);
                        ih.SendModKey(Interceptor.Keys.LeftShift, Interceptor.Keys.Right);
                    }
                }
                if (bg.CancellationPending)
                {
                    break;
                }
                if (i < csv.Count())
                {
                    //NewLine();
                    //ih.SendKey(Interceptor.Keys.Enter);
                    ih.SendModKey(Interceptor.Keys.LeftShift, Interceptor.Keys.Down);
                    for (int j = 0; j < line.Count()-1; j++)
                    {
                        //ih.SendKey(Interceptor.Keys.Down);
                        ih.SendModKey(Interceptor.Keys.LeftShift, Interceptor.Keys.Left);
                    }
                    
                }
            }

            ih.UnloadDriver();
        }

        private void SendKey(char c)
        {
            // SendKeys.Send(c.ToString());
            short b = Convert.ToSByte(c);
            // ((Keyboard.ScanCodeShort)b).ToString();
            // Keyboard.VirtualKeyShort vKB = ((Keyboard.VirtualKeyShort)b);
            // kb.SendVirtual(vKB);//Keyboard.ScanCodeShort.KEY_0);
            //In_sim.Keyboard.TextEntry(c);
            //In_sim.Keyboard.Sleep(strokeDelay);
           
           
            Thread.Sleep(strokeDelay);
        }
        private void NewCell()
        {
            //SendKeys.Send("{TAB}");
            //kb.SendVirtual(Keyboard.VirtualKeyShort.TAB);
            //Thread.Sleep(strokeDelay);
        }
        private void NewLine()
        {
            // SendKeys.Send("{ENTER}");
           // kb.SendVirtual(Keyboard.VirtualKeyShort.RETURN);
            //Thread.Sleep(strokeDelay);
        }
       
    }
}
