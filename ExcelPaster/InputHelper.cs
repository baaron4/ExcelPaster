using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using Interceptor;
using System.Threading;

namespace ExcelPaster
{
    public class InputHelper
    {
        public void test()
        {
            Input input = new Input();

            // Be sure to set your keyboard filter to be able to capture key presses and simulate key presses
            // KeyboardFilterMode.All captures all events; 'Down' only captures presses for non-special keys; 'Up' only captures releases for non-special keys; 'E0' and 'E1' capture presses/releases for special keys
            input.KeyboardFilterMode = KeyboardFilterMode.All;
            // You can set a MouseFilterMode as well, but you don't need to set a MouseFilterMode to simulate mouse clicks

            // Finally, load the driver
            input.Load();

            
            input.SendKeys(Keys.Enter);  // Presses the ENTER key down and then up (this constitutes a key press)
            Thread.Sleep(1);
            input.SendText("A");
            input.Unload();
        }
        
    }
}
