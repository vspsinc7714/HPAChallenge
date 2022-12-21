using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics.Contracts;
using System.Windows.Forms;
using HPACodingChallenge.UINavigation;
using System.Threading;

namespace HPACodingChallenge
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Retrieving notepad process Id to retrieve UIAutomation elements.");
                ManageNotePad notePadMgr = new ManageNotePad();

                //Click the "File" menu
               // notePadMgr.ClickFileMenuItem();
               // Console.WriteLine("Click File menu item.");

                // Thread.Sleep(2000);

                //Click the "New" menu item
                notePadMgr.ClickNewMenuItem();
                Console.WriteLine("Click New menu item.");

                Thread.Sleep(2000);

                //Enter "Hello World"
                notePadMgr.AddText("Hello World");
                Console.WriteLine("Add text.");

                Thread.Sleep(2000);

                //Click the "Save" menu item
                notePadMgr.ClickSave();
                Console.WriteLine("Click Save");

                Thread.Sleep(2000);

                //Click the "Save As" menu item
                notePadMgr.ClickSaveAs();
                Console.WriteLine("Click Save as.");

             //   Thread.Sleep(2000);

                //get user input
                Console.WriteLine("Press any key to quit");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exiting applicaton. " + ex.Message);
            }
            
        }
    }
}
