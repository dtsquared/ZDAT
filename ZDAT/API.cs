using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ZDAT
{
    public class API
    {
        public void SendText(IntPtr hWnd, string Text)
        {
            int i = 0;

            do
            {
                if (i % 4 == 0)
                {
                    PInvoke.MH.sendString(hWnd, Text);
                }
                PInvoke.MH.sendString(hWnd, Text);
                System.Threading.Thread.Sleep(250);
                i += 1;
            } while (PInvoke.MH.GetWindowTextRaw(hWnd) != Text);
        }

        public void PopUpBox(string Title, string Button)
        {
            IntPtr PhWnd = PInvoke.Win.GetHandle(Title);
            List<IntPtr> chld;
            int x = 0;

            if (PhWnd != IntPtr.Zero)
            {
                chld = PInvoke.Win.GetChildWindows(PhWnd);

                for (int i = 0; i < chld.Count; i++)
                {
                    if (PInvoke.MH.GetWindowTextRaw(chld[i]) == Button)
                    {
                        PInvoke.Mouse.LeftClick(chld[i]);
                        do
                        {
                            if (x % 8 == 0)
                            {
                                PInvoke.Mouse.LeftClick(chld[i]);
                            }
                            Thread.Sleep(250);
                            x += 1;
                        } while (PInvoke.Win.GetHandle(Title) != IntPtr.Zero);
                        break;
                    }
                }
            }
        }

        public void SendKey(IntPtr hWnd, Keys key)
        {
            PInvoke.MH.sendKey(hWnd, key, true);
        }
    }


    [XmlRootAttribute("ProgProfile", Namespace = "ZDAT", IsNullable = false)]
    public class Item
    {
    }
}
