using System;
using System.Collections.Generic;
using System.Threading;
using PInvoke;
using System.Diagnostics;


namespace ZDAT
{
    public abstract class Task
    {
        public Process StartZD(string BranchNumber, string ZDPath, string UserName)
        {
            Process ZD = new Process();
            ZD.StartInfo.FileName = ZDPath;
            ZD.StartInfo.Arguments = "B" + BranchNumber + ":5632 WESNETZD";
            ZD.Start();

            return ZD;
        }

        public void HideZDWindow()
        {
            IntPtr MainhWnd = Win.GetHandle("3615 - WESNET ZD");
            List<IntPtr> childhWnd = Win.GetChildWindows(MainhWnd);

            Win.HideWindow(childhWnd[8]);
            Win.HideWindow(childhWnd[4]);
        }

        public bool DialogHandler(string WinTitle, string ButtonText)
        {
            IntPtr hWnd = Win.GetHandle(WinTitle);
            List<IntPtr> childhWnd;
            int iCounter = 0;

            if (hWnd != IntPtr.Zero)
            {
                Console.WriteLine("Dialog Window " + WinTitle + " found!");

                childhWnd = Win.GetChildWindows(hWnd);

                for (int i = 0; i < childhWnd.Count; i++)
                {
                    if (MH.GetWindowTextRaw(childhWnd[i]) == ButtonText)
                    {
                        Console.WriteLine("Button " + ButtonText + " found!");

                        Mouse.LeftClick(childhWnd[i]);

                        while (Win.GetHandle(WinTitle) != IntPtr.Zero)
                        {
                            iCounter += 1;
                            Thread.Sleep(500);

                            if (iCounter % 20 == 0)
                            {
                                Console.WriteLine("Attempting to click " + ButtonText + " again.");
                                Mouse.LeftClick(childhWnd[i]);
                            }
                        }
                    }
                }
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
