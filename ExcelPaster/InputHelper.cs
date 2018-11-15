using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using Interceptor;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;

namespace ExcelPaster
{
    public class InputHelper
    {
        private Input input = null;

        public void SendKey(char c)
        {
            if (!input.IsLoaded)
            {
                LoadDriver();
            }
            if (input.IsLoaded)
            {
                if (c == '_')
                {
                    input.SendKey(Interceptor.Keys.LeftShift, KeyState.Down);
                    Thread.Sleep(5);
                    input.SendKey(Interceptor.Keys.DashUnderscore);
                    Thread.Sleep(5);
                    input.SendKey(Interceptor.Keys.LeftShift, KeyState.Up);
                    Thread.Sleep(5);
                    
                } else
                if (Char.IsUpper(c))
                {
                    input.SendKey(Interceptor.Keys.LeftShift, KeyState.Down);
                    Thread.Sleep(5);
                    input.SendText(c.ToString());
                    Thread.Sleep(5);
                    input.SendKey(Interceptor.Keys.LeftShift, KeyState.Up);
                    Thread.Sleep(5);
                    
                }
                else
                {
                    input.SendText(c.ToString());
                    Thread.Sleep(5);
                   
                }
               
            }
           
        }
        public void SendKeys(string s)
        {
            if (!input.IsLoaded)
            {
                LoadDriver();
            }
            if (input.IsLoaded)
            {
                input.SendText(s);
                Thread.Sleep(5);

            }
        }
        public void SendKey(Interceptor.Keys k)
        {
            if (!input.IsLoaded)
            {
                LoadDriver();
            }
            if (input.IsLoaded)
            {
                input.SendKey(k, KeyState.Down);
                Thread.Sleep(5);
                input.SendKey(k, KeyState.Up);
                Thread.Sleep(5);

            }
            
        }
        public void SendModKey(Interceptor.Keys m, Interceptor.Keys k)
        {
            if (!input.IsLoaded)
            {
                LoadDriver();
            }
            if (input.IsLoaded)
            {
                input.SendKey(m, KeyState.Down);
                Thread.Sleep(5);
                input.SendKey(k, KeyState.Down);
                Thread.Sleep(5);
                input.SendKey(k, KeyState.Up);
                Thread.Sleep(5);
                input.SendKey(m, KeyState.Up);
                Thread.Sleep(5);

            }
        }
        public  void LoadDriver()
        {
            if (input == null)
            {
                input = new Input();
                // Be sure to set your keyboard filter to be able to capture key presses and simulate key presses
                // KeyboardFilterMode.All captures all events; 'Down' only captures presses for non-special keys; 'Up' only captures releases for non-special keys; 'E0' and 'E1' capture presses/releases for special keys
                input.KeyPressDelay = 0;

                input.KeyboardFilterMode = KeyboardFilterMode.All;
                // You can set a MouseFilterMode as well, but you don't need to set a MouseFilterMode to simulate mouse clicks

                // Finally, load the driver
                input.Load();
                //wait for user input to gather key driver
            }
        }
        public void UnloadDriver()
        {
            if (input != null)
            {
                input.Unload();
            }
        }
        public void test()
        {
            if (input == null)
            {
                input = new Input();
                // Be sure to set your keyboard filter to be able to capture key presses and simulate key presses
                // KeyboardFilterMode.All captures all events; 'Down' only captures presses for non-special keys; 'Up' only captures releases for non-special keys; 'E0' and 'E1' capture presses/releases for special keys
                input.KeyPressDelay = 0;

                input.KeyboardFilterMode = KeyboardFilterMode.All;
                // You can set a MouseFilterMode as well, but you don't need to set a MouseFilterMode to simulate mouse clicks

                // Finally, load the driver
                input.Load();
            }

            if (input.IsLoaded)
            {
                //MessageBox.Show("Press enter!");
                //input.SendKey(Interceptor.Keys.S, KeyState.Down);
                //Thread.Sleep(50);
                //input.SendKey(Interceptor.Keys.S, KeyState.Up);
                //Thread.Sleep(50);
                //input.SendKeys(Interceptor.Keys.A);  // Presses the ENTER key down and then up (this constitutes a key press)
                Thread.Sleep(1000);
                input.SendText("Bravofuckyoumodabucker");
            }

            //input.Unload();
        }
        
    }
}
