using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RPA_Tech.Models
{
    public class Notepad
    {
        public Notepad()
        {
            NotepadProcess = getNotepad();
            NotepadTextbox = getNotepadTextBox();
        }
        public Process NotepadProcess { get; set; }
        public IntPtr NotepadTextbox { get; set; }
        public IntPtr SaveAsWindowInput { get; set; }

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public Process getNotepad()
        {
            Process[] localAll = Process.GetProcesses();
            Process notepad = new Process();

            Process[] processes = Process.GetProcessesByName("notepad");
            if (processes.Length == 0)
            {
                Console.WriteLine("Notepad not running...Starting.");
                notepad = Process.Start("notepad");
                Thread.Sleep(500);
            }
            else
            {
                Console.WriteLine("Notepad found running.");
                notepad = processes[0];

            }
            return notepad;

        }

        public IntPtr getSaveAsTextBox()
        {
            IntPtr saveAsWindow = FindWindow(null, "Save As");
            IntPtr saveAsPanel1 = FindWindowEx(saveAsWindow, IntPtr.Zero, "DUIViewWndClassName", null);
            IntPtr saveAsPanel2 = FindWindowEx(saveAsPanel1, IntPtr.Zero, "DirectUIHWND", null);
            IntPtr saveAsPanel3 = FindWindowEx(saveAsPanel2, IntPtr.Zero, "FloatNotifySink", null);
            IntPtr saveAsPanel4 = FindWindowEx(saveAsPanel3, IntPtr.Zero, "ComboBox", null);
            IntPtr saveAsEdit = FindWindowEx(saveAsPanel4, IntPtr.Zero, "Edit", null);

            return saveAsEdit;
        }

        public IntPtr getConfirmWindow()
        {
            return FindWindow(null, "Confirm Save As");
        }

        public IntPtr getNotepadTextBox()
        {
            return FindWindowEx(NotepadProcess.MainWindowHandle, IntPtr.Zero, "Edit", null);
        }

        public void saveNotepad()
        {
            Console.WriteLine("Saving notepad.");
            SetForegroundWindow(NotepadProcess.MainWindowHandle);
            SendKeys.SendWait("%{f}");
            SendKeys.SendWait("+{s}");
            SendKeys.SendWait("^+{s}");
        }

        public void inputNotepadText(string notepadTextInput)
        {
            Console.WriteLine("Sending text to notepad");

            const int WM_SETTEXT = 0X000C;
            SendMessage(NotepadTextbox, WM_SETTEXT, 0, notepadTextInput);
        }

        public void saveFile(string fileNameToSave)
        {
            Console.WriteLine("Saving file as: " + "'" + fileNameToSave + "'");

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileNameWithPath = desktopPath + @"\" + fileNameToSave + ".txt";

            const int WM_SETTEXT = 0X000C;
            SendMessage(SaveAsWindowInput, WM_SETTEXT, 0, fileNameWithPath);
            SendKeys.SendWait("%{s}");
            IntPtr confirmWindow = FindWindow(null, "Confirm Save As");
            if (confirmWindow != IntPtr.Zero)
            {
                Console.WriteLine("File already exists... Overwriting.");
                SendKeys.SendWait("%{y}");
            }
        }
    }
}
