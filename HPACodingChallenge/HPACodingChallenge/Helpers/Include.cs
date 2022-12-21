using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace HPACodingChallenge.Helpers
{
    internal static class Include
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(IntPtr hwnd, string name);

        //This is a replacement for Cursor.Position in WinForms
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        public static void BringWindowToFront(IntPtr hwnd)
        {
            IntPtr rtrn = SetForegroundWindow(hwnd);
        }

        public static IntPtr LocateWindow(IntPtr hwnd, string windowName)
        {
            return FindWindow(hwnd, windowName);
        }


        public static Process FindNotePad(string appRunnigProcessName, string appNotRunnigProcessName)
        {
            Process notePadProcess;
            //see if notepad is currently running
            Process[] process = Process.GetProcessesByName(appRunnigProcessName);
            Console.WriteLine("Looking for currently running notepad process.");

            //start process if not currently running
            if (process.Length == 0)
            {
                Console.WriteLine("Notepad was not running. Starting notepad process.");
                notePadProcess = Process.Start(appNotRunnigProcessName);
            }
            else
            {
                notePadProcess = process[0];
            }

            return notePadProcess;
        }

        public static void SendFileMenuShortCutKeys(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("%(f)");
        }
        public static void SendNewMenuShortCutKeys(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("^(n)");
        }
        public static void SendSaveMenuShortCutKeys(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("^(s)");
        }

        public static void SendSaveAsMenuShortCutKeys(IntPtr hwnd)
        {
            IntPtr rtrn = SetForegroundWindow(hwnd);
            SendKeys.SendWait("(^+)s");
        }

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }
    }
}
