using HPACodingChallenge.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using System.Drawing;
using System.Windows.Input;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using System.Data.Common;
using System.Net;
using System.Windows.Media.Media3D;
using System.Windows.Media;
using System.Windows;
using Point = System.Windows.Point;

namespace HPACodingChallenge.UINavigation
{
    internal class ManageNotePad
    {
        public enum UIView
        {
            Raw,
            Control,
            Content
        }

        string processNameNotRunning = Path.Combine(Environment.SystemDirectory, "notepad.exe");
        string processNameRunning = "notepad";
        string fileName = Path.Combine("HPAChallenge");

        //public string fileName { get; set; }
        Process notePadProcess { get; set; }

        IntPtr mainWindowHandle;

        AutomationElement root;
        List<AutomationElement> mainWindowElements;
        List<AutomationElement> fileMenuElements;
        List<AutomationElement> saveDialogElements;

        public ManageNotePad()
        {
            try
            {
                mainWindowElements = new List<AutomationElement>();
                fileMenuElements = new List<AutomationElement>();

                //find notepad and set element root
                notePadProcess = Include.FindNotePad(processNameRunning, processNameNotRunning);
                root = FindElementByProcessId(notePadProcess.Id);
                //walk main window tree
                WalkControlElements(root, ref mainWindowElements, UIView.Raw);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Fatal error in Constructor: " + ex.ToString());
                throw;
            }
        
        }

        public void AddText(string newText)
        {
            try
            {
                //sanity checks
                if (String.IsNullOrEmpty(newText)) return;

                //running Windows 11
                AutomationElement textBlock = mainWindowElements.FirstOrDefault(e => e.Current.ControlType.LocalizedControlType.Equals("document") && e.Current.ClassName.Equals("RichEditD2DPT"));
               
                //try Windows 10
                if (textBlock == null)
                {
                    textBlock = mainWindowElements.FirstOrDefault(e => e.Current.ControlType.LocalizedControlType.Equals("document") && e.Current.ClassName.Equals("Edit"));
                }

                if ((textBlock == null) || (textBlock.Current.IsEnabled == false))
                {
                    Console.WriteLine("Unable to locate Notepad input area to add text.");
                    return;
                }

                if (textBlock.Current.IsKeyboardFocusable == false)
                {
                    textBlock.SetFocus();
                }

                object valuePattern;
                if (!textBlock.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
                {

                    Thread.Sleep(100);
                    SendKeys.SendWait(newText);
                }
                else
                {
                    ((ValuePattern)valuePattern).SetValue(newText);
                }
            } 
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        

        public void ClickFileMenuItem()
        {
            try
            {
                //make sure notpad has focus
                mainWindowHandle = new IntPtr(root.Current.NativeWindowHandle);
                Include.BringWindowToFront(mainWindowHandle);

                //display file menu
                AutomationElement fileMenuItem = mainWindowElements.FirstOrDefault(e => e.Current.ControlType.LocalizedControlType.Equals("menu item") && e.Current.Name.Equals("File"));
                //System.Drawing.Point p = new System.Drawing.Point(Convert.ToInt32(fileMenuItem.Current.BoundingRectangle.X-5), Convert.ToInt32(fileMenuItem.Current.BoundingRectangle.Y-5));
                //MouseClickElement(p);
                Include.SendFileMenuShortCutKeys(mainWindowHandle);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }            
        }

        public void ClickNewMenuItem()
        {
            try
            {
                //make sure notepad has focus
                AddText("");
                mainWindowHandle = new IntPtr(root.Current.NativeWindowHandle);
                Include.SendNewMenuShortCutKeys(mainWindowHandle);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public void ClickSave()
        {
            try
            {
                mainWindowHandle = new IntPtr(root.Current.NativeWindowHandle);
                Include.SendSaveMenuShortCutKeys(mainWindowHandle);
                FileSave("Save as");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public void ClickSaveAs()
        {
            try
            {
                AddText("");
                Include.SendSaveAsMenuShortCutKeys(mainWindowHandle);
                FileSave("Save as");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }            
        }

        private void FileSave(string windowName)
        {
            try
            {
                //find Save dialog to make sure it's there
                IntPtr winHandle = Include.LocateWindow(IntPtr.Zero, windowName);
                //click Save button
                if (winHandle == IntPtr.Zero)
                {
                    Console.WriteLine("Unable to locate Save Dialog");
                    return;
                }
                saveDialogElements = new List<AutomationElement>();
                AutomationElement element = AutomationElement.FromHandle(winHandle);
                //walk dailog tree to find save button, etc.
                WalkControlElements(element, ref saveDialogElements, UIView.Raw);
                AutomationElement comboBox = saveDialogElements.FirstOrDefault(b => b.Current.LocalizedControlType.Equals("combo box") && b.Current.Name.StartsWith("File name:"));
                AddFileName(comboBox, fileName);
                AutomationElement saveButton = saveDialogElements.FirstOrDefault(b => b.Current.LocalizedControlType.Equals("button") && b.Current.Name.StartsWith("Save"));
                if (saveButton == null)
                {
                    Console.WriteLine("Unable to locate Save Dialog, retry.");
                    if (element.Current.IsKeyboardFocusable == false)
                    {
                        element.SetFocus();
                    }
                    //walk dailog tree to find save button, etc.
                    saveDialogElements = new List<AutomationElement>();
                    WalkControlElements(element, ref saveDialogElements, UIView.Raw);
                    comboBox = saveDialogElements.FirstOrDefault(b => b.Current.LocalizedControlType.Equals("combo box") && b.Current.Name.StartsWith("File name:"));
                    AddFileName(comboBox, fileName);
                    saveButton = saveDialogElements.FirstOrDefault(b => b.Current.LocalizedControlType.Equals("button") && b.Current.Name.StartsWith("Save"));
                }

                if (saveButton == null)
                {
                    Console.WriteLine("Unable to locate Save Dialog.");
                    return;
                }
                ClickElement(saveButton);

                //check for confirm messagebox
                winHandle = Include.LocateWindow(IntPtr.Zero, "Confirm Save As");
                if (winHandle == IntPtr.Zero)
                {
                    return;   //we are done
                }

                List<AutomationElement> confirmDialogElements = new List<AutomationElement>(); 
                AutomationElement confirmDialog = AutomationElement.FromHandle(winHandle);
                WalkControlElements(confirmDialog, ref confirmDialogElements, UIView.Control);
                saveButton = confirmDialogElements.FirstOrDefault(b => b.Current.LocalizedControlType.Equals("button") && b.Current.Name.StartsWith("Yes"));
                ClickElement(saveButton);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }            
        }

        public void AddFileName(AutomationElement elem, string newText)
        {
            try
            {
                //sanity checks
                if (String.IsNullOrEmpty(newText)) return;
              
                if ((elem == null) || (elem.Current.IsEnabled == false))
                {
                    Console.WriteLine("Unable to locate file name textbox for update.");
                    return;
                }

                if (elem.Current.IsKeyboardFocusable == false)
                {
                    elem.SetFocus();
                }

                object valuePattern;
                if (!elem.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
                {

                    Thread.Sleep(100);
                    SendKeys.SendWait(newText);
                }
                else
                {
                    ((ValuePattern)valuePattern).SetValue(newText);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static AutomationElement FindElementByProcessId(int pid)
        {
            int cntr = 0;
            AutomationElement element = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.ProcessIdProperty, pid);

            while ((element == null) && (++cntr < 15000))
            {
                element = AutomationElement.RootElement.FindFirst(TreeScope.Children, conds);
            }
            return element;
        }

        private void WalkControlElements(AutomationElement rootElement, ref List<AutomationElement> elementsList, UIView viewCondition)
        {
            try
            {
                TreeWalker walker = null;

                if (viewCondition == UIView.Raw)
                    walker = TreeWalker.RawViewWalker;

                if (viewCondition == UIView.Control)
                    walker = TreeWalker.ControlViewWalker;

                if (viewCondition == UIView.Content)
                    walker = TreeWalker.ContentViewWalker;

                AutomationElement elementNode = walker.GetFirstChild(rootElement);

                while (elementNode != null)
                {
                    //string strTemp = "ControlType: " + elementNode.Current.ControlType.LocalizedControlType + "; ClassName: " + elementNode.Current.ClassName.ToString() + "; Name: " + elementNode.Current.Name;
                    //Console.WriteLine(strTemp);
                    elementsList.Add(elementNode);
                    WalkControlElements(elementNode, ref elementsList, viewCondition);
                    elementNode = walker.GetNextSibling(elementNode);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }            
        }

        private void ClickElement(AutomationElement autoElement)
        {
            try
            {
                var invokePattern = autoElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern?.Invoke();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
        }

        private void MouseClickElement(System.Drawing.Point point)
        {
            Include.LeftMouseClick(point.X, point.Y);
        }
    }
}
