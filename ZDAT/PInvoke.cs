using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;


namespace PInvoke
{
    public class Win
    {
        #region DLLImports

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, ShowWindowCommands nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //Converts the first pixel of a programs rendering area to a point on the screen
        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hwnd, ref Point lpPoint);

        //Converts the screen coordinates of a specified point on the screen to client-area coordinates.
        [DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hwnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern bool PtInRegion(IntPtr hrgn, int X, int Y);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        #endregion

        #region User Defined Types
        enum ShowWindowCommands : int
        {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            Hide = 0,
            /// <summary>
            /// Activates and displays a window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when displaying the window
            /// for the first time.
            /// </summary>
            Normal = 1,
            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            ShowMinimized = 2,
            /// <summary>
            /// Maximizes the specified window.
            /// </summary>
            Maximize = 3, // is this the right value?
            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>      
            ShowMaximized = 3,
            /// <summary>
            /// Displays a window in its most recent size and position. This value
            /// is similar to <see cref="Win32.ShowWindowCommand.Normal"/>, except
            /// the window is not activated.
            /// </summary>
            ShowNoActivate = 4,
            /// <summary>
            /// Activates the window and displays it in its current size and position.
            /// </summary>
            Show = 5,
            /// <summary>
            /// Minimizes the specified window and activates the next top-level
            /// window in the Z order.
            /// </summary>
            Minimize = 6,
            /// <summary>
            /// Displays the window as a minimized window. This value is similar to
            /// <see cref="Win32.ShowWindowCommand.ShowMinimized"/>, except the
            /// window is not activated.
            /// </summary>
            ShowMinNoActive = 7,
            /// <summary>
            /// Displays the window in its current size and position. This value is
            /// similar to <see cref="Win32.ShowWindowCommand.Show"/>, except the
            /// window is not activated.
            /// </summary>
            ShowNA = 8,
            /// <summary>
            /// Activates and displays the window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when restoring a minimized window.
            /// </summary>
            Restore = 9,
            /// <summary>
            /// Sets the show state based on the SW_* value specified in the
            /// STARTUPINFO structure passed to the CreateProcess function by the
            /// program that started the application.
            /// </summary>
            ShowDefault = 10,
            /// <summary>
            ///  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread
            /// that owns the window is not responding. This flag should only be
            /// used when minimizing windows from a different thread.
            /// </summary>
            ForceMinimize = 11
        }
        #endregion

        #region WindowStyles
        //Window Style Params
        const int GWL_ID = -12;
        const int GWL_STYLE = -16;
        const int GWL_EXSTYLE = -20;
        #endregion

        public static bool IsOverlapped(IntPtr Handle)
        {
            int Hidden = (int)0x4C00000;
            bool result = (GetWindowLong(Handle, GWL_STYLE) == Hidden);

            return result;
        }

        public static void HideWindow(IntPtr hWnd)
        {
            ShowWindow(hWnd, ShowWindowCommands.Hide);
        }

        public static bool IsChecked(IntPtr Handle)
        {
            int Checked = (int)0x58010004;
            bool result = (GetWindowLong(Handle, GWL_STYLE) == Checked);
            Console.WriteLine("Window Style Value: " + GetWindowLong(Handle, GWL_STYLE));

            return result;
        }

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }

        /// <summary>
        /// Delegate for the EnumChildWindows method
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="parameter">Caller-defined variable; we use it for a pointer to our list</param>
        /// <returns>True to continue enumerating, false to bail.</returns>
        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        public static Point ScrnToCli(IntPtr handle)
        {
            Point point = Cursor.Position;
            bool foundwindow = ScreenToClient(handle, ref point);

            if (foundwindow)
            {
                return point;
            }
            else
            {
                return Point.Empty;
            }
        }

        /// <summary>
        ///    Translates a point in an applications client area to a point
        ///    on the screen.
        /// </summary>
        /// <param name="handle">The applications client area handle.</param>
        /// <param name="point">A point relative to origin of the client area.</param>
        /// <returns>
        ///    <c>Point</c> if the operation succeeded, <c>Threw ArgumentOutOfRangeException</c> otherwise.
        /// </returns>
        public static Point ClientScreen(IntPtr handle, Point point)
        {
            bool found = ClientToScreen(handle, ref point);
            if (found)
            {
                return point;
            }
            else
            {
                return new Point(0, 0);
            }
        }

        public static IntPtr GetHandle(string winTitle)
        {
            IntPtr handle = FindWindow(default(string), winTitle);
            Process[] processlist = Process.GetProcesses();

            if (handle == IntPtr.Zero)
            {
                for (int i = 0; i <= processlist.Length - 1; i++)
                {
                    if (processlist[i].MainWindowTitle.Contains(winTitle))
                    {
                        handle = processlist[i].MainWindowHandle;
                    }
                    else
                    {
                        handle = IntPtr.Zero;
                    }
                }
            }

            return handle;
        }

        public static void BringToTop(IntPtr handle)
        {
            bool result = SetForegroundWindow(handle);
        }
    }

    public class Mouse
    {
        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr WindowFromPoint(Point point);

        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010,
        }

        private const int BM_CLICK = 0x00F5;
        public static void LeftClick(IntPtr handle)
        {
            SendMessage(handle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
        }
    }

    public class MH
    {

        [Flags]
        public enum LB
        {
            /* Multi-Select Only */
            /// <summary>
            ///   Selects an item in a multiple-selection list box and, if necessary, scrolls the item into view
            ///   wParam  Specifies how to set the selection. If this parameter is TRUE, the item is selected and highlighted;
            ///   if it is FALSE, the highlight is removed and the item is no longer selected.
            ///   lParam Specifies the zero-based index of the item to set. If this parameter is –1, 
            ///   the selection is added to or removed from all items, depending on the value of wParam, and no scrolling occurs.
            ///   Return Value-If an error occurs, the return value is LB_ERR.
            ///   Remarks-Use this message only with multiple-selection list boxes.
            /// </summary>
            LB_SETSEL = 0x0185,

            ///<summary>
            ///  Fills a buffer with an array of integers that specify the item numbers of selected items in a multiple-selection list box
            ///  wParam- The maximum number of selected items whose item numbers are to be placed in the buffer.
            ///  lParam-A pointer to a buffer large enough for the number of integers specified by the wParam parameter.
            ///  Return-The return value is the number of items placed in the buffer. If the list box is a single-selection list box, 
            ///  the return value is LB_ERR.
            ///</summary>
            LB_GETSELITEMS = 0x0191,

            /// <summary>
            ///   Gets the total number of selected items in a multiple-selection list box
            ///   wParam and lParam-Not used; must be zero.
            ///   Return Value-The return value is the count of selected items in the list box.
            ///   If the list box is a single-selection list box, the return value is LB_ERR.
            /// </summary>
            LB_GETSELCOUNT = 0x0190,

            /* Single Select Only */
            /// <summary>
            ///   Gets the index of the currently selected item, if any, in a single-selection list box
            ///   wParam and lParam Not used; must be zero.
            ///   Return Value-In a single-selection list box, the return value is the zero-based index of the currently selected item.
            ///   If there is no selection, the return value is LB_ERR.
            /// </summary>
            LB_GETCURSEL = 0x0188,

            /// <summary>
            ///   wParam The zero-based index of the item before the first item to be searched.
            ///   When the search reaches the bottom of the list box, 
            ///   it continues from the top of the list box back to the item specified by the wParam parameter. 
            ///   If wParam is –1, the entire list box is searched from the beginning.
            ///   lParam A pointer to the null-terminated string that contains the prefix for which to search.
            ///   The search is case independent return Value If the search is successful,
            ///   the return value is the index of the selected item.
            ///   If the search is unsuccessful, LB_ERR use this message only with single-selection list boxes.
            ///   You cannot use it to set or remove a selection in a multiple-selection list box
            /// </summary>
            LB_SELECTSTRING = 0x018C,

            /* Either */
            ///<summary>
            ///  Selects a string and scrolls it into view, if necessary.
            ///  When the new string is selected, the list box removes the highlight from the previously selected string.
            ///  wParam Specifies the zero-based index of the string that is selected. 
            ///  If this parameter is -1, the list box is set to have no selection.
            ///  lParam not used
            ///</summary>
            LB_SETCURSEL = 0x0186,

            ///<summary>
            ///  Finds the first string in a list box that begins with the specified string
            ///  wParam The zero-based index of the item before the first item to be searched. 
            ///  When the search reaches the bottom of the list box, it continues searching from 
            ///  the top of the list box back to the item specified by the wParam parameter. 
            ///  If wParam is –1, the entire list box is searched from the beginning.
            ///  lParam A pointer to the null-terminated string that contains the string for which to search.
            ///  The search is case independent, so this string can contain any combination of uppercase and lowercase letters.
            ///  Return Value-The return value is the index of the matching item, or LB_ERR if the search was unsuccessful.
            ///</summary>
            LB_FINDSTRING = 0x018F,

            /// <summary>
            ///   Gets the number of items in a list box
            ///   wParam and lParam-zero
            ///   The return value is the number of items in the list box, or LB_ERR if an error occurs
            /// </summary>
            LB_GETCOUNT = 0x018B,

            ///<summary>
            ///  Gets the selection state of an item
            ///  wParam The zero-based index of the item
            ///  lParam This parameter is not used
            ///  Return Value-If an item is selected, the return value is greater than zero; otherwise, it is zero.
            ///  If an error occurs, the return value is LB_ERR.
            ///</summary>
            LB_GETSEL = 0x0187,

            ///<summary>
            ///  Gets a string from a list box.
            ///  wParam The zero-based index of the string to retrieve.
            ///  lParam A pointer to the buffer that will receive the string;
            ///  The buffer must have sufficient space for the string and a terminating null character.
            ///  Return Value length of the string, in TCHARs, excluding the terminating null character.
            ///</summary>
            LB_GETTEXT = 0x0189,

            /// <summary>
            ///   Removes all items from a list box
            /// </summary>
            LB_RESETCONTENT = 0x0184,

            /// <summary>
            /// Sets the width, in pixels, by which a list box can be scrolled horizontally
            /// (the scrollable width). If the width of the list box is smaller than this
            /// value, the horizontal scroll bar horizontally scrolls items in the list box.
            /// If the width of the list box is equal to or greater than this value, the
            /// horizontal scroll bar is hidden.
            /// </summary>
            LB_SETHORIZONTALEXTENT = 0x0194,

            /// <summary>
            /// Gets the width, in pixels, that a list box can be scrolled horizontally
            /// (the scrollable width) if the list box has a horizontal scroll bar.
            /// </summary>
            LB_GETHORIZONTALEXTENT = 0x0193,

            /// <summary>
            /// Gets the index of the first visible item in a list box. Initially the item
            /// with index 0 is at the top of the list box, but if the list box contents have
            /// been scrolled another item may be at the top. The first visible item in a
            /// multiple-column list box is the top-left item.
            /// </summary>
            LB_GETTOPINDEX = 0x018E,

            /// <summary>
            /// Ensures that the specified item in a list box is visible.
            /// </summary>
            LB_SETTOPINDEX = 0x0197,

            /// <summary>
            /// Inserts a string or item data into a list box. Unlike the LB_ADDSTRING
            /// message, the LB_INSERTSTRING message does not cause a list with the
            /// LBS_SORT style to be sorted.
            /// </summary>
            LB_INSERTSTRING = 0x0181,

            /// <summary>
            /// Deletes a string in a list box.
            /// </summary>
            LB_DELETESTRING = 0x0182,

            /// <summary>
            /// Gets the application-defined value associated with the specified list box item.
            /// </summary>
            LB_GETITEMDATA = 0x0199
        }

        //[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        //private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);

        //Listbox
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        //GetWindowTextRaw
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, [Out] StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public const int WM_USER = 0x400;
        public const int WM_COPYDATA = 0x004A;
        public const int WM_SETTEXT = 0x000C;
        public const int WM_GETTEXT = 0x000D;
        public const uint WM_GETTEXTLENGTH = 0x000E;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_SETFOCUS = 0x7;
        ///<summary>
        ///  Fills a buffer with an array of integers that specify the item numbers of selected items in a multiple-selection list box
        ///  wParam- The maximum number of selected items whose item numbers are to be placed in the buffer.
        ///  lParam-A pointer to a buffer large enough for the number of integers specified by the wParam parameter.
        ///  Return-The return value is the number of items placed in the buffer. If the list box is a single-selection list box, 
        ///  the return value is LB_ERR.
        ///</summary>
        public const int LB_GETSELITEMS = 0x0191;

        /// <summary>
        ///   Gets the number of items in a list box
        ///   wParam and lParam-zero
        ///   The return value is the number of items in the list box, or LB_ERR if an error occurs
        /// </summary>
        public const int LB_GETCOUNT = 0x018B;

        /// <summary>
        /// Gets the application-defined value associated with the specified list box item.
        /// </summary>
        public const int LB_GETITEMDATA = 0x0199;

        ///<summary>
        ///  Gets a string from a list box.
        ///  wParam The zero-based index of the string to retrieve.
        ///  lParam A pointer to the buffer that will receive the string;
        ///  The buffer must have sufficient space for the string and a terminating null character.
        ///  Return Value length of the string, in TCHARs, excluding the terminating null character.
        ///</summary>
        public const int LB_GETTEXT = 0x0189;

        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public Point pt;
        }

        public static int sendString(IntPtr hWnd, string lParam)
        {
            StringBuilder sb = new StringBuilder(lParam);
            int result = SendMessage(hWnd, WM_SETTEXT, (IntPtr)sb.Length, sb.ToString());
            return result;
        }

        public static void sendKey(IntPtr hwnd, Keys key, bool extended)
        {
            uint keyCode = (uint)key;
            uint scanCode = MapVirtualKey((uint)keyCode, 0);
            uint lParam;

            lParam = (0x00000001 | (scanCode << 16));
            if (extended)
            {
                lParam = lParam | 0x01000000;
            }

            //KEY DOWN
            PostMessage(hwnd, (UInt32)WM_KEYDOWN, (IntPtr)keyCode, (IntPtr)lParam);

            //KEY UP
            PostMessage(hwnd, (UInt32)WM_KEYUP, (IntPtr)keyCode, (IntPtr)lParam);
        }

        public static void setFocus(IntPtr hwnd)
        {
            SendMessage(hwnd, WM_SETFOCUS, 0, 0);
        }

        public static string GetWindowTextRaw(IntPtr hwnd)
        {
            int length = (int)SendMessage(hwnd, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
            StringBuilder sb = new StringBuilder(length + 1);
            SendMessage(hwnd, WM_GETTEXT, (IntPtr)sb.Capacity, sb);

            return sb.ToString();
        }

        public static List<string> GetListboxItems(IntPtr hwnd)
        {
            List<string> list = new List<string>();
            int itemCount = SendMessage(hwnd, LB_GETCOUNT, IntPtr.Zero, IntPtr.Zero);
            StringBuilder sb = new StringBuilder(50);

            for (int i = 0; i < itemCount; i++)
            {
                SendMessage(hwnd, LB_GETTEXT, new IntPtr(i), sb);
                list.Add(sb.ToString());
            }

            return list;
        }
    }
}
