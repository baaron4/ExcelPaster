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
        public void test()
        {
            if (input == null)
            {
                input = new Input();
                input.Load();
            }
            

            // Be sure to set your keyboard filter to be able to capture key presses and simulate key presses
            // KeyboardFilterMode.All captures all events; 'Down' only captures presses for non-special keys; 'Up' only captures releases for non-special keys; 'E0' and 'E1' capture presses/releases for special keys
            input.KeyPressDelay = 0;
            
            input.KeyboardFilterMode = KeyboardFilterMode.All;
            // You can set a MouseFilterMode as well, but you don't need to set a MouseFilterMode to simulate mouse clicks

            // Finally, load the driver


           
            if (input.IsLoaded)
            {
               // MessageBox.Show("Press enter!");
                input.SendKey(Interceptor.Keys.S, KeyState.Down);
                Thread.Sleep(50);
                input.SendKey(Interceptor.Keys.S, KeyState.Up);
                Thread.Sleep(50);
                //input.SendKeys(Interceptor.Keys.A);  // Presses the ENTER key down and then up (this constitutes a key press)
                //Thread.Sleep(1000);
                //input.SendText("Bravofuckyoumodabucker");
            }
            
            //input.Unload();
        }
        
    }
}
