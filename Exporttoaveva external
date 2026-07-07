// ============================================================================
//  ExportToAveva_External.cs
//  ---------------------------------------------------------------------------
//  Standalone .EXE version of the "export_test" Aveva/PDMS export macro.
//  Runs from OUTSIDE Tekla against an ALREADY-OPEN Tekla 2024 model.
//
//  How it connects:
//    Tekla Open API uses .NET Remoting. When Tekla + a model are open,
//    "new Model()" attaches to that running instance from any process.
//
//  Build settings (Tekla 2024):
//    - Project Type   : Windows Forms App (.NET Framework)  or Console App
//    - Target Framework: .NET Framework 4.8   (required for Tekla 2024)
//    - Platform Target : x64                  (NOT Any CPU)
//    - NuGet          : Tekla.Structures 2024.0.0
//                       Tekla.Structures.Model 2024.0.0
//
//  Namespace reference (valid for 2024/2025):
//    https://developer.tekla.com/doc/tekla-structures/2025/tekla-structures-45473
//
//  NOTE: No "$" string interpolation is used anywhere (avoids CS1056).
//        All strings are built with "+" concatenation.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Tekla.Structures;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
using TSMUI = Tekla.Structures.Model.UI;
using TSG = Tekla.Structures.Geometry3d;

namespace ExportToAvevaExternal
{
    public class Program
    {
        // -- Win32 API ---------------------------------------------------------
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool EnumChildWindows(IntPtr hWndParent,
                                            EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg,
                                         IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, int dx, int dy,
                                       uint dwData, IntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

        // -- Constants ---------------------------------------------------------
        const uint BM_CLICK = 0x00F5;
        const uint WM_SETFOCUS = 0x0007;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const int  SW_RESTORE = 9;

        const byte VK_CONTROL = 0x11;
        const byte VK_A = 0x41;
        const uint KEYEVENTF_KEYUP = 0x0002;

        delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int Left, Top, Right, Bottom; }

        // -- Logging -----------------------------------------------------------
        static readonly string LogPath = @"C:\TeklaStructures\logs\export_test.log";

        static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg + Environment.NewLine);
            }
            catch { /* never let logging break the run */ }
        }

        // ====================================================================
        //  EXTERNAL ENTRY POINT
        //  Replaces the old "public static void Run(IScript akit)".
        //  akit was never used in the original body, so nothing is lost.
        // ====================================================================
        [STAThread]
        static void Main(string[] args)
        {
            Log("=== export_test (EXTERNAL) START ===");

            // -- 1. Verify Tekla + model are open (avoids RemotingException) --
            Model model = new Model();
            if (!model.GetConnectionStatus() ||
                model.GetInfo().ModelPath == string.Empty)
            {
                string warn = "Tekla Structures is not running, or no model is open." +
                              Environment.NewLine +
                              "Please open the model first, then run this exporter.";
                Log("ERROR: " + warn);
                MessageBox.Show(warn, "Connection Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Log("Connected to model: " + model.GetInfo().ModelPath);

            try
            {
                RunExport();
            }
            catch (Exception ex)
            {
                Log("FATAL: " + ex.Message);
                MessageBox.Show("Export failed. Original error: " + ex.Message,
                    "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Log("=== export_test (EXTERNAL) END ===");
        }

        // ====================================================================
        //  Core export routine (unchanged logic from the original macro)
        // ====================================================================
        static void RunExport()
        {
            // XSDATADIR is read from the open model via .NET Remoting.
            string xsEnv = "";
            TeklaStructuresSettings.GetAdvancedOption("XSDATADIR", ref xsEnv);
            string EnginePath = xsEnv
                + @"\environments\common\extensions\ExportToAveva\ExportToSoftware.exe";

            Log("EnginePath = " + EnginePath);

            if (!File.Exists(EnginePath))
            {
                Log("ERROR: engine not found, aborting.");
                MessageBox.Show("Export engine not found:" + Environment.NewLine + EnginePath,
                    "Missing Engine", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Kill any stale export window from a previous run so FindWindow
            // does not latch onto an old, empty instance.
            foreach (Process p in Process.GetProcessesByName("ExportToSoftware"))
            {
                try { p.Kill(); Log("Killed stale ExportToSoftware.exe (PID " + p.Id + ")"); }
                catch { }
            }
            Thread.Sleep(1000);

            Process proc = new Process();
            proc.StartInfo.FileName = EnginePath;
            proc.StartInfo.Arguments = "PDMS";

            try { proc.Start(); Log("Started ExportToSoftware.exe PDMS"); }
            catch (Exception ex) { Log("ERROR starting engine: " + ex.Message); return; }

            // -- Poll up to 20s for the Aveva window (10 x 2s) -----------------
            IntPtr hMainWnd = IntPtr.Zero;
            for (int i = 0; i < 10; i++)
            {
                hMainWnd = FindAvevaWindow();
                if (hMainWnd != IntPtr.Zero) { Log("Found Aveva window after " + (i * 2) + "s, hWnd=" + hMainWnd); break; }
                Thread.Sleep(2000);
            }
            if (hMainWnd == IntPtr.Zero) { Log("ERROR: Aveva window never appeared."); return; }

            ShowWindow(hMainWnd, SW_RESTORE);
            SetForegroundWindow(hMainWnd);
            Thread.Sleep(1500);

            // -- Poll up to 10s for the grid (5 x 2s) --------------------------
            IntPtr hGrid = IntPtr.Zero;
            for (int i = 0; i < 5; i++)
            {
                hGrid = FindLargestWindow8Child(hMainWnd);
                if (hGrid != IntPtr.Zero) { Log("Found grid after " + (i * 2) + "s, hWnd=" + hGrid); break; }
                Thread.Sleep(2000);
            }
            if (hGrid == IntPtr.Zero) { Log("ERROR: grid control not found."); return; }

            // Give the grid extra time to load all rows before selecting.
            Thread.Sleep(5000);

            // -- Focus the grid with a real click ------------------------------
            RECT gridRect;
            GetWindowRect(hGrid, out gridRect);
            int clickX = (gridRect.Left + gridRect.Right) / 2;
            int clickY = (gridRect.Top + gridRect.Bottom) / 2;
            Log("Grid rect L=" + gridRect.Left + " T=" + gridRect.Top +
                " R=" + gridRect.Right + " B=" + gridRect.Bottom +
                "  click=(" + clickX + "," + clickY + ")");

            SetForegroundWindow(hMainWnd);
            Thread.Sleep(200);
            SetCursorPos(clickX, clickY);
            Thread.Sleep(150);
            mouse_event(MOUSEEVENTF_LEFTDOWN, clickX, clickY, 0, IntPtr.Zero);
            Thread.Sleep(80);
            mouse_event(MOUSEEVENTF_LEFTUP, clickX, clickY, 0, IntPtr.Zero);
            Thread.Sleep(500);

            // -- Select all rows: low-level Ctrl+A, then SendKeys as fallback --
            Log("Sending Ctrl+A via keybd_event");
            keybd_event(VK_CONTROL, 0, 0, IntPtr.Zero);
            keybd_event(VK_A, 0, 0, IntPtr.Zero);
            Thread.Sleep(60);
            keybd_event(VK_A, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
            Thread.Sleep(400);

            Log("Sending Ctrl+A via SendKeys (fallback)");
            try { SendKeys.SendWait("^a"); } catch (Exception ex) { Log("SendKeys failed: " + ex.Message); }
            Thread.Sleep(600);

            // -- Click Export --------------------------------------------------
            IntPtr hExportBtn = FindButtonByText(hMainWnd, "Export");
            if (hExportBtn == IntPtr.Zero) { Log("ERROR: Export button not found."); return; }
            Log("Clicking Export, hWnd=" + hExportBtn);
            SendMessage(hExportBtn, BM_CLICK, IntPtr.Zero, IntPtr.Zero);

            // -- Run save.cs right after Export (works via Remoting) -----------
            Log("Running save.cs");
            Tekla.Structures.Model.Operations.Operation.RunMacro("save.cs");

            // -- Wait for export, then Close -----------------------------------
            Thread.Sleep(5000);
            hMainWnd = FindAvevaWindow();
            if (hMainWnd != IntPtr.Zero)
            {
                IntPtr hCloseBtn = FindButtonByText(hMainWnd, "Close");
                if (hCloseBtn != IntPtr.Zero) { Log("Clicking Close"); SendMessage(hCloseBtn, BM_CLICK, IntPtr.Zero, IntPtr.Zero); }
                else Log("Close button not found.");
            }
        }

        // -- Find Aveva window -- exact title then partial match ---------------
        static IntPtr FindAvevaWindow()
        {
            IntPtr result = IntPtr.Zero;
            string[] titles = new string[]
            {
                "Export to Aveva : 3.5",
                "Export to Aveva",
                "ExportToSoftware"
            };
            foreach (string title in titles)
            {
                result = FindWindow(null, title);
                if (result != IntPtr.Zero) return result;
            }
            EnumChildWindows(GetDesktopWindow(), delegate (IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                if (sb.ToString().IndexOf("Export to Aveva",
                    StringComparison.OrdinalIgnoreCase) >= 0 && IsWindowVisible(hWnd))
                {
                    result = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        // -- Find largest WindowsForms10.Window.8 child by screen area ---------
        static IntPtr FindLargestWindow8Child(IntPtr hParent)
        {
            IntPtr bestHandle = IntPtr.Zero;
            int bestArea = 0;
            EnumChildWindows(hParent, delegate (IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder cls = new StringBuilder(256);
                GetClassName(hWnd, cls, 256);
                string clsStr = cls.ToString();
                if (clsStr.StartsWith("WindowsForms10.Window.8",
                    StringComparison.OrdinalIgnoreCase))
                {
                    StringBuilder txt = new StringBuilder(128);
                    GetWindowText(hWnd, txt, 128);
                    if (txt.ToString().Trim().Length > 0) return true;
                    RECT r;
                    if (GetWindowRect(hWnd, out r))
                    {
                        int area = (r.Right - r.Left) * (r.Bottom - r.Top);
                        if (area > bestArea) { bestArea = area; bestHandle = hWnd; }
                    }
                }
                return true;
            }, IntPtr.Zero);
            return bestHandle;
        }

        // -- Find BUTTON child by exact text -----------------------------------
        static IntPtr FindButtonByText(IntPtr hParent, string buttonText)
        {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(hParent, delegate (IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder cls = new StringBuilder(64);
                GetClassName(hWnd, cls, 64);
                if (cls.ToString().IndexOf("BUTTON", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    StringBuilder txt = new StringBuilder(128);
                    GetWindowText(hWnd, txt, 128);
                    if (txt.ToString().Trim().Equals(buttonText,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        found = hWnd;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }
    }
}
