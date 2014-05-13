using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Threading;
using System.Windows.Forms;
using PInvoke;
using System.ComponentModel;
using System.Windows.Threading;
using System.Diagnostics;


namespace ZDAT
{
    public class Automation
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

        public void NewOrderEntry(Process ZD)
        {

        }

        public bool EnterOrders(IEnumerable<OpenXMLReader.Order> Orders, int TotalLines, bool EDIEnabled, string Branch, BackgroundWorker worker, DoWorkEventArgs e)
        {
            #region Declarations
            IntPtr OPhandle;
            List<IntPtr> OPchildhWnd;
            IntPtr OHhandle;
            List<IntPtr> OHchildhWnd;
            IntPtr Phandle;
            List<IntPtr> PchildhWnd;
            List<string> MissedLines = new List<string>();

            int LineNumber = 0;
            int percentComplete = 0;
            int iCounter = 1;
            string lastPO = "";
            bool newOrder = false;
            #endregion

            #region Load_OP
            OPhandle = Win.GetHandle(Branch + " - Sales Order Management - New");
            OPchildhWnd = Win.GetChildWindows(OPhandle);
            #endregion

            #region Load_OH
            OHhandle = Win.GetHandle(Branch + " - Sales Order Management - Header Information");
            if (OHhandle == IntPtr.Zero)
            {
                OHhandle = Win.GetHandle(Branch + " - Sales Order Management - Header Information - New");
            }
            OHchildhWnd = Win.GetChildWindows(OHhandle);
            #endregion

            Win.BringToTop(OPhandle);

            if (OPhandle != IntPtr.Zero)
            {
                #region Get_First_PO
                IEnumerator<OpenXMLReader.Order> fistPO = Orders.GetEnumerator();
                fistPO.MoveNext();
                lastPO = fistPO.Current.PONumber;
                #endregion

                foreach (OpenXMLReader.Order o in Orders)
                {
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        #region Order_Header
                        #region Fill_Order_Header
                        if (o.PONumber != lastPO)
                        {
                            Mouse.LeftClick(OPchildhWnd[(int)OP.OrderHeader]);

                            do
                            {
                                if (iCounter % 25 == 0)
                                {
                                    Mouse.LeftClick(OPchildhWnd[(int)OP.OrderHeader]);
                                }
                                Thread.Sleep(300);
                                iCounter += 1;
                            } while (Win.IsOverlapped(OHhandle) == true);

                            iCounter = 1;
                            MH.setFocus(OHchildhWnd[(int)OH.CustomerPO]);
                            Thread.Sleep(500);
                            MH.sendString(OHchildhWnd[(int)OH.CustomerPO], lastPO);
                            Thread.Sleep(500);
                            if (EDIEnabled == true)
                            {
                                MH.setFocus(OHchildhWnd[(int)OH.EdiPO]);
                                Thread.Sleep(250);
                                MH.sendString(OHchildhWnd[(int)OH.EdiPO], lastPO);
                                Thread.Sleep(250);
                                MH.sendKey(OHchildhWnd[(int)OH.EdiPO], Keys.Tab, true);
                                Thread.Sleep(300);
                            }

                            Mouse.LeftClick(OHchildhWnd[(int)OH.Order]);
                            Thread.Sleep(300);
                            Mouse.LeftClick(OHchildhWnd[(int)OH.Order]);
                        #endregion Fill_Order_Header

                            #region Save_Order
                            Console.Write("Saving Order");
                            Phandle = IntPtr.Zero;
                            Mouse.LeftClick(OHchildhWnd[(int)OH.Save]);
                            do
                            {
                                if (iCounter % 25 == 0)
                                {
                                    Mouse.LeftClick(OHchildhWnd[(int)OH.Save]);
                                }
                                Thread.Sleep(250);
                                Phandle = Win.GetHandle("OEPAD");
                                iCounter += 1;
                            } while (Phandle == IntPtr.Zero);
                            iCounter = 1;

                            PchildhWnd = Win.GetChildWindows(Phandle);
                            for (int i = 0; i < PchildhWnd.Count; i++)
                            {
                                if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                {
                                    Mouse.LeftClick(PchildhWnd[i]);
                                    do
                                    {
                                        if (iCounter % 4 == 0)
                                        {
                                            Mouse.LeftClick(PchildhWnd[i]);
                                        }
                                        Thread.Sleep(250);
                                        iCounter += 1;
                                    } while (Win.GetHandle("OEPAD") != IntPtr.Zero);
                                }
                            }
                            Console.WriteLine("Save complete!");
                            Console.WriteLine();
                            Phandle = IntPtr.Zero;
                            iCounter = 1;

                            #region Print_Order
                            Console.Write("Printing ticket");
                            for (int i = 0; i <= 20; i++)
                            {
                                Phandle = Win.GetHandle(Branch + " - Printer Selection");
                                Thread.Sleep(250);
                                if (Phandle != IntPtr.Zero)
                                    break;
                            }

                            PchildhWnd = Win.GetChildWindows(Phandle);
                            for (int i = 0; i < PchildhWnd.Count; i++)
                            {
                                if (MH.GetWindowTextRaw(PchildhWnd[i]) == "Se&lect")
                                {
                                    Mouse.LeftClick(PchildhWnd[i]);
                                    do
                                    {
                                        if (iCounter % 25 == 0)
                                        {
                                            Mouse.LeftClick(PchildhWnd[i]);
                                        }
                                        Thread.Sleep(250);
                                        iCounter += 1;
                                    } while (Win.GetHandle(Branch + " - Printer Selection") != IntPtr.Zero);
                                }
                            }
                            Console.WriteLine();
                            Phandle = IntPtr.Zero;
                            iCounter = 1;

                            Console.Write("Confirming ticket printed");
                            do
                            {
                                if (iCounter % 25 == 0)
                                {
                                    break;
                                }
                                Phandle = Win.GetHandle("O111C");
                                Thread.Sleep(250);
                                iCounter += 1;
                            } while (Phandle == IntPtr.Zero);

                            if (iCounter % 25 != 0)
                            {
                                PchildhWnd = Win.GetChildWindows(Phandle);
                                for (int i = 0; i < PchildhWnd.Count; i++)
                                {
                                    if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                    {
                                        Mouse.LeftClick(PchildhWnd[i]);
                                        do
                                        {
                                            if (iCounter % 25 == 0)
                                            {
                                                Mouse.LeftClick(PchildhWnd[i]);
                                            }
                                            Thread.Sleep(250);
                                            iCounter += 1;
                                        } while (Win.GetHandle("O111C") != IntPtr.Zero);
                                    }
                                }
                            }
                            Console.WriteLine();
                            Phandle = IntPtr.Zero;
                            iCounter = 1;
                            #endregion
                            #endregion

                            #region Clear_Order
                            ////Console.Write("\r\nClearing order");
                            //Mouse.LeftClick(OHchildhWnd[(int)OH.Clear]);
                            //do
                            //{
                            //    //Console.Write(".");
                            //    if (iCounter % 25 == 0)
                            //    {
                            //        Mouse.LeftClick(OHchildhWnd[(int)OH.Clear]);
                            //    }
                            //    Phandle = Window.GetHandle("OEPAD");
                            //    Thread.Sleep(250);
                            //    iCounter += 1;
                            //} while (Phandle == IntPtr.Zero);
                            //iCounter = 1;
                            //PchildhWnd = Window.GetChildWindows(Phandle);
                            //for (int i = 0; i < PchildhWnd.Count; i++)
                            //{
                            //    if (MH.GetWindowTextRaw(PchildhWnd[i]) == "&Yes")
                            //    {
                            //        Mouse.LeftClick(PchildhWnd[i]);
                            //        do
                            //        {
                            //            if (iCounter % 25 == 0)
                            //            {
                            //                Mouse.LeftClick(PchildhWnd[i]);
                            //            }
                            //            Thread.Sleep(250);
                            //            iCounter += 1;
                            //        } while (Window.GetHandle("OEPAD") != IntPtr.Zero);
                            //    }
                            //}
                            //iCounter = 1;
                            //Phandle = IntPtr.Zero;
                            ////Console.WriteLine();
                            #endregion

                            #region Lost_Sale
                            ////Console.WriteLine("Entering reason for lost sale.");

                            //do
                            //{
                            //    Phandle = Window.GetHandle(Branch + " - Lost Sale");
                            //    Thread.Sleep(250);
                            //    iCounter += 1;
                            //} while (Phandle == IntPtr.Zero);

                            //PchildhWnd = Window.GetChildWindows(Phandle);

                            //for (int i = 0; i < PchildhWnd.Count; i++)
                            //{
                            //    if (MH.GetWindowTextRaw(PchildhWnd[i]) == "")
                            //    {
                            //        MH.setFocus(PchildhWnd[i]);
                            //        MH.sendString(PchildhWnd[i], "K");
                            //        do
                            //        {
                            //            if (iCounter % 25 == 0)
                            //            {
                            //                MH.sendString(PchildhWnd[i], "K");
                            //            }
                            //            Thread.Sleep(250);
                            //            iCounter += 1;
                            //        } while (Window.GetHandle("OEPAD") != IntPtr.Zero);
                            //    }
                            //}
                            //iCounter = 1;
                            //Phandle = IntPtr.Zero;
                            #endregion


                            Mouse.LeftClick(OHchildhWnd[(int)OH.OrderPad]);
                            do
                            {
                                if (iCounter % 20 == 0)
                                {
                                    Win.BringToTop(OPhandle);
                                    Mouse.LeftClick(OHchildhWnd[(int)OH.OrderPad]);
                                }
                                Thread.Sleep(250);
                                iCounter += 1;
                            } while (Win.IsOverlapped(OHhandle) == false);
                            iCounter = 1;
                        }
                        #endregion Order_Header

                        LineNumber += 1;
                        MainWindow.cqSOM.Enqueue(LineNumber);

                        percentComplete = LineNumber * 100;
                        percentComplete = percentComplete / TotalLines;
                        worker.ReportProgress(percentComplete);

                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                        }
                        else
                        {
                            #region Customer
                            if (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.Customer]) == "")
                            {
                                do
                                {
                                    MH.sendString(OPchildhWnd[(int)OP.Customer], o.Customer);
                                    MH.sendKey(OPchildhWnd[(int)OP.Customer], Keys.Enter, true);
                                    lastPO = o.PONumber;
                                    newOrder = true;
                                }
                                while (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.Customer]) != o.Customer);

                                do
                                {
                                    if (iCounter % 10 == 0)
                                    {
                                        MH.setFocus(OPchildhWnd[(int)OP.Customer]);
                                        MH.sendKey(OPchildhWnd[(int)OP.Customer], Keys.Enter, true);
                                    }
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                } while (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.ShipToDesc]) == "");

                                if (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.LineIncr]) != "2")
                                {
                                    MH.sendString(OPchildhWnd[(int)OP.LineIncr], "2");
                                }
                                iCounter = 1;
                            }

                            Thread.Sleep(250);

                            #endregion

                            #region Ship_To
                            if (o.Area.ToString() != MH.GetWindowTextRaw(OPchildhWnd[(int)OP.ShipToNum]).ToString())
                            {
                                MH.setFocus(OPchildhWnd[(int)OP.ShipToNum]); 
                                MH.sendString(OPchildhWnd[(int)OP.ShipToNum], o.Area);
                                MH.setFocus(OPchildhWnd[(int)OP.ShipToNum]);
                                MH.sendKey(OPchildhWnd[(int)OP.ShipToNum], Keys.Enter, true);
                                Thread.Sleep(250);

                                do
                                {
                                    if (iCounter % 25 == 0)
                                    {
                                        MH.sendKey(OPchildhWnd[(int)OP.ShipToNum], Keys.Enter, true);
                                    }
                                    Phandle = Win.GetHandle("OEPAD");
                                    iCounter += 1;
                                } while (Phandle == IntPtr.Zero);

                                PchildhWnd = Win.GetChildWindows(Phandle);
                                for (int i = 0; i < PchildhWnd.Count; i++)
                                {
                                    if (MH.GetWindowTextRaw(PchildhWnd[i]) == "&No")
                                    {
                                        Mouse.LeftClick(PchildhWnd[i]);
                                        do
                                        {
                                            if (iCounter % 25 == 0)
                                            {
                                                Mouse.LeftClick(PchildhWnd[i]);
                                            }
                                            Thread.Sleep(250);
                                            iCounter += 1;
                                        } while (Win.GetHandle("OEPAD") != IntPtr.Zero);
                                    }
                                }
                                iCounter = 1;
                                Phandle = IntPtr.Zero;
                            }
                            #endregion

                            #region Qty
                            Console.WriteLine("Entering Qty.");
                            MH.setFocus(OPchildhWnd[(int)OP.Qty]);
                            MH.sendString(OPchildhWnd[(int)OP.Qty], o.Qty);
                            MH.sendKey(OPchildhWnd[(int)OP.Qty], Keys.Enter, true);
                            Thread.Sleep(250);
                            #endregion

                            if (worker.CancellationPending)
                            {
                                e.Cancel = true;
                            }
                            else
                            {
                                #region Product


                                Console.WriteLine("Entering part number " + o.Part);
                                Win.BringToTop(OPhandle);

                                Thread.Sleep(250);
                                if (newOrder == true)
                                {
                                    MH.setFocus(OPchildhWnd[(int)OP.Product]);
                                    Thread.Sleep(300);
                                    MH.sendString(OPchildhWnd[(int)OP.Product], o.Part.PadLeft(11, '0'));
                                    Thread.Sleep(250);
                                    MH.sendKey(OPchildhWnd[(int)OP.Product], Keys.Enter, true);
                                    newOrder = false;
                                }
                                else
                                {
                                    if (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.ProdCode]) == "")
                                    {
                                        MH.sendString(OPchildhWnd[(int)OP.Product], o.Part);
                                    }
                                    MH.sendString(OPchildhWnd[(int)OP.Product], o.Part);
                                    MH.sendKey(OPchildhWnd[(int)OP.Product], Keys.Tab, true);
                                }
                                Console.Write("Waiting for product to load");
                                iCounter = 1;

                                do
                                {
                                    Console.Write(".");
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                    if (iCounter % 20 == 0)
                                    {
                                        MH.setFocus(OPchildhWnd[(int)OP.Product]);
                                        Thread.Sleep(300);
                                        if (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.ProdCode]) == "")
                                        {
                                            MH.sendString(OPchildhWnd[(int)OP.Product], o.Part);
                                        }
                                        MH.sendKey(OPchildhWnd[(int)OP.Product], Keys.Enter, true);
                                    }
                                    if (Win.GetHandle(Branch + "  -  Price and Availability") != IntPtr.Zero)
                                    {
                                        Phandle = Win.GetHandle(Branch + "  -  Price and Availability");
                                        PchildhWnd = Win.GetChildWindows(Phandle);
                                        for (int i = 0; i < PchildhWnd.Count; i++)
                                        {
                                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "&Close")
                                            {
                                                Win.BringToTop(OPhandle);
                                                Mouse.LeftClick(PchildhWnd[i]);
                                                do
                                                {
                                                    Thread.Sleep(250);
                                                } while (Win.GetHandle(Branch + "  -  Price and Availability") != IntPtr.Zero);
                                                MissedLines.Add("LN: " + LineNumber + " Qty: " + o.Qty + " SIM: " + o.Part);
                                                break;
                                            }
                                        }
                                    }

                                    Phandle = Win.GetHandle("OEPAD Error");
                                    if (Phandle != IntPtr.Zero)
                                    {
                                        PchildhWnd = Win.GetChildWindows(Phandle);
                                        for (int i = 0; i < PchildhWnd.Count; i++)
                                        {
                                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                            {
                                                Mouse.LeftClick(PchildhWnd[i]);
                                            }
                                        }
                                    }
                                } while (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.Desc]) == "");
                                Phandle = IntPtr.Zero;
                                iCounter = 1;
                                Console.WriteLine();
                                #endregion

                                #region Price
                                //Mouse.LeftClick(childhWnd[(int)OP.Price]);

                                #region Price Overrides
                                ////////////////////
                                //                                                                          //
                                // Used for price overrides - Tested ~140 line items with 100% success rate //
                                //                                                                          //
                                ////////////////////

                                //do
                                //{
                                //    //Console.WriteLine("Waiting for price to be set..");
                                //    MH.sendWMString(childhWnd[(int)OP.Price], MH.WM_SETTEXT, "0.001");
                                //    Double.TryParse(MH.GetWindowTextRaw(childhWnd[(int)OP.Price]), out Price);
                                //    Thread.Sleep(250);
                                //} while (Price <= 0.000);
                                #endregion

                                //MH.sendKey(childhWnd[(int)OP.Price], Keys.Enter, true);
                                //Thread.Sleep(500);
                                #endregion

                                #region Add
                                Console.WriteLine("Adding to order.");
                                Thread.Sleep(500);
                                Mouse.LeftClick(OPchildhWnd[(int)OP.Add]);
                                Console.Write("Waiting for application to enter a ready state");
                                do
                                {
                                    Console.Write(".");

                                    if (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.Add]) == "Upd&ate")
                                    {
                                        Mouse.LeftClick(OPchildhWnd[(int)OP.NewItem]);
                                        MissedLines.Add("LN: " + LineNumber + " Qty: " + o.Qty + " SIM: " + o.Part);
                                        Thread.Sleep(2000);
                                    }
                                    if (iCounter % 25 == 0)
                                    {
                                        Mouse.LeftClick(OPchildhWnd[(int)OP.Add]);
                                    }
                                    Phandle = Win.GetHandle("OEPAD Error");
                                    if (Phandle != IntPtr.Zero)
                                    {
                                        PchildhWnd = Win.GetChildWindows(Phandle);
                                        for (int i = 0; i < PchildhWnd.Count; i++)
                                        {
                                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                            {
                                                Mouse.LeftClick(PchildhWnd[i]);
                                            }
                                        }
                                    }
                                    Thread.Sleep(500);
                                    iCounter += 1;
                                } while (MH.GetWindowTextRaw(OPchildhWnd[(int)OP.Qty]) != "0");

                                Console.WriteLine("\r\n");
                                iCounter = 1;
                                #endregion

                                Thread.Sleep(500);
                                Phandle = Win.GetHandle("OEPAD Error");
                                if (Phandle != IntPtr.Zero)
                                {
                                    PchildhWnd = Win.GetChildWindows(Phandle);
                                    for (int i = 0; i < PchildhWnd.Count; i++)
                                    {
                                        if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                        {
                                            Mouse.LeftClick(PchildhWnd[i]);
                                        }
                                    }
                                }
                                lastPO = o.PONumber;
                            }
                        }
                    }//if cancel != true
                } //foreach order

                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                }
                else
                {
                    Win.BringToTop(OPhandle);
                    Mouse.LeftClick(OPchildhWnd[(int)OP.OrderHeader]);

                    do
                    {
                        if (iCounter % 25 == 0)
                        {
                            Mouse.LeftClick(OPchildhWnd[(int)OP.OrderHeader]);
                        }
                        Thread.Sleep(250);
                        iCounter += 1;
                    } while (Win.IsOverlapped(OHhandle) == true);
                    iCounter = 1;

                    MH.setFocus(OHchildhWnd[(int)OH.CustomerPO]);
                    Thread.Sleep(500);
                    MH.sendString(OHchildhWnd[(int)OH.CustomerPO], lastPO);
                    Thread.Sleep(500);
                    if (EDIEnabled == true)
                    {
                        MH.setFocus(OHchildhWnd[(int)OH.EdiPO]);
                        Thread.Sleep(250);
                        MH.sendString(OHchildhWnd[(int)OH.EdiPO], lastPO);
                        Thread.Sleep(250);
                        MH.sendKey(OHchildhWnd[(int)OH.EdiPO], Keys.Tab, true);
                        Thread.Sleep(300);
                    }

                    Mouse.LeftClick(OHchildhWnd[(int)OH.Order]);

                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        #region Save_Order
                        Console.Write("Saving Order");
                        Phandle = IntPtr.Zero;
                        Mouse.LeftClick(OHchildhWnd[(int)OH.Save]);
                        do
                        {
                            if (iCounter % 25 == 0)
                            {
                                Mouse.LeftClick(OHchildhWnd[(int)OH.Save]);
                            }
                            Thread.Sleep(250);
                            Phandle = Win.GetHandle("OEPAD");
                            iCounter += 1;
                        } while (Phandle == IntPtr.Zero);
                        iCounter = 1;

                        PchildhWnd = Win.GetChildWindows(Phandle);
                        for (int i = 0; i < PchildhWnd.Count; i++)
                        {
                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                            {
                                Mouse.LeftClick(PchildhWnd[i]);
                                do
                                {
                                    if (iCounter % 4 == 0)
                                    {
                                        Mouse.LeftClick(PchildhWnd[i]);
                                    }
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                } while (Win.GetHandle("OEPAD") != IntPtr.Zero);
                            }
                        }
                        Console.WriteLine();
                        Phandle = IntPtr.Zero;
                        iCounter = 1;
                        #endregion

                        #region Print_Order
                        iCounter = 1;
                        Console.Write("Printing ticket");
                        do
                        {
                            Phandle = Win.GetHandle(Branch + " - Printer Selection");
                            Thread.Sleep(250);
                            if (iCounter % 10 == 0)
                            {
                                break;
                            }
                            iCounter++;
                        } while (Phandle == IntPtr.Zero);

                        iCounter = 1;
                        PchildhWnd = Win.GetChildWindows(Phandle);
                        for (int i = 0; i < PchildhWnd.Count; i++)
                        {
                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "Se&lect")
                            {
                                Mouse.LeftClick(PchildhWnd[i]);
                                do
                                {
                                    if (iCounter % 25 == 0)
                                    {
                                        Mouse.LeftClick(PchildhWnd[i]);
                                    }
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                } while (Win.GetHandle(Branch + " - Printer Selection") != IntPtr.Zero);
                            }
                        }
                        Console.WriteLine();
                        Phandle = IntPtr.Zero;

                        iCounter = 1;
                        Console.Write("Confirming ticket printed");
                        do
                        {
                            Phandle = Win.GetHandle("O111C");
                            Thread.Sleep(250);

                            if (iCounter % 25 == 0)
                            {
                                break;
                            }
                            iCounter++;
                        } while (Phandle == IntPtr.Zero);

                        iCounter = 1;
                        PchildhWnd = Win.GetChildWindows(Phandle);
                        for (int i = 0; i < PchildhWnd.Count; i++)
                        {
                            if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                            {
                                Mouse.LeftClick(PchildhWnd[i]);
                                do
                                {
                                    if (iCounter % 25 == 0)
                                    {
                                        Mouse.LeftClick(PchildhWnd[i]);
                                    }
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                } while (Win.GetHandle("O111C") != IntPtr.Zero);
                            }
                        }
                        Console.WriteLine();
                        Phandle = IntPtr.Zero;
                        iCounter = 1;
                        #endregion

                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                        }
                        else
                        {
                            Win.BringToTop(OHhandle);
                            Mouse.LeftClick(OHchildhWnd[(int)OH.OrderPad]);
                            do
                            {
                                if (iCounter % 15 == 0)
                                {
                                    Win.BringToTop(OHhandle);
                                    Thread.Sleep(250);
                                    Mouse.LeftClick(OPchildhWnd[(int)OH.OrderPad]);
                                }
                                Thread.Sleep(250);
                                iCounter += 1;
                            } while (Win.IsOverlapped(OHhandle) == false);

                            Win.BringToTop(OPhandle);
                            iCounter = 1;
                        }
                    }
                }
            }//if window is found
            else
            {
                MessageBox.Show("SOM window not found!");
                return false;
            }
            return true;
        } //Main

        public bool ConvertEDI(MainWindow.CustomerRange custRange, BackgroundWorker worker, DoWorkEventArgs e)
        {
            #region Declarations
            IntPtr hWnd = Win.GetHandle("3615 - Order Search Selection Criteria");
            List<IntPtr> childhWnd = new List<IntPtr>();

            IntPtr phWnd;
            List<IntPtr> pchildhWnd = new List<IntPtr>();

            List<SalesOrder> orderList = new List<SalesOrder>();
            SalesOrder salesOrder = new SalesOrder();

            int iCounter = 0;
            bool dialogResult;
            #endregion

            if (hWnd != IntPtr.Zero)
            {
                Win.BringToTop(hWnd);
                childhWnd = Win.GetChildWindows(hWnd);

                #region Setup_OS
                Mouse.LeftClick(childhWnd[(int)OS.itmNoneOfAbove]);
                Thread.Sleep(250);

                Mouse.LeftClick(childhWnd[(int)OS.ordsInquiries]);
                Thread.Sleep(250);

                MH.sendString(childhWnd[(int)OS.fltrCustFrom], custRange.CustFrom);
                Thread.Sleep(250);

                MH.sendString(childhWnd[(int)OS.fltrCustTo], custRange.CustTo);
                Thread.Sleep(250);
                #endregion

                Mouse.LeftClick(childhWnd[(int)OS.Search]);
                Thread.Sleep(500);

                dialogResult = DialogHandler("LUORDSEL", "OK");

                do
                {
                    Thread.Sleep(500);
                    if (Win.GetHandle("LUORDSEL") != IntPtr.Zero)
                    {
                        phWnd = Win.GetHandle("LUORDSEL");
                        pchildhWnd = Win.GetChildWindows(phWnd);

                        for (int i = 0; i < pchildhWnd.Count; i++)
                        {
                            if (MH.GetWindowTextRaw(pchildhWnd[i]) == "OK")
                            {
                                Mouse.LeftClick(pchildhWnd[i]);
                                //TODO:
                                //STOP OR MOVE TO NEXT DPC
                                //BECAUSE NO INQUIRIES WERE FOUND
                            }
                        }
                    }
                } while (Win.GetHandle("3615 - Order Search Results") == IntPtr.Zero);

                Thread.Sleep(1000); //Give the window time to finish loading

                hWnd = Win.GetHandle("3615 - Order Search Results");
                childhWnd = Win.GetChildWindows(hWnd);

                do
                {
                    salesOrder.ContractNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.ContractNo]);
                    salesOrder.CreatedBy = MH.GetWindowTextRaw(childhWnd[(int)OSR.CreatedBy]);
                    salesOrder.Customer = MH.GetWindowTextRaw(childhWnd[(int)OSR.Customer]);
                    salesOrder.CustomerPO = MH.GetWindowTextRaw(childhWnd[(int)OSR.CustPO]);
                    salesOrder.CustPartNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.CustPartNum]);
                    salesOrder.OrderNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.OrderNum]);
                    salesOrder.OrderStatus = MH.GetWindowTextRaw(childhWnd[(int)OSR.OrderStatus]);

                    if (salesOrder.OrderStatus != "Open")
                    {
                        MH.sendKey(childhWnd[(int)OSR.OrderGrid], Keys.Down, true);
                        do
                        {
                            Thread.Sleep(500);
                        } while (MH.GetWindowTextRaw(childhWnd[(int)OSR.OrderStatus]) == "");
                    }
                } while (salesOrder.OrderStatus != "Open");

                //Click manage order twice due to an issue with focus
                //even when focus is set first clicking once does not work
                Mouse.LeftClick(childhWnd[(int)OSR.ManageOrder]);

                do
                {
                    Thread.Sleep(500);
                    iCounter += 1;

                    if (Win.GetHandle("3615 - Order Search Results") != hWnd)
                    {
                        DialogHandler("3615 - Order Search Results", "OK");
                    }

                    DialogHandler("LUORDRSLT Error", "OK");

                    if (iCounter % 10 == 0)
                    {
                        Mouse.LeftClick(childhWnd[(int)OSR.ManageOrder]);
                    }

                } while (Win.GetHandle("3615 - Sales Order Management - " + salesOrder.OrderNumber) == IntPtr.Zero);
                iCounter = 0;

                hWnd = Win.GetHandle("3615 - Sales Order Management - " + salesOrder.OrderNumber);
                childhWnd = Win.GetChildWindows(hWnd);

                do
                {
                    Thread.Sleep(1000);
                } while (MH.GetWindowTextRaw(childhWnd[(int)OP.Customer]) == "");

                do
                {
                    Thread.Sleep(500);

                    iCounter += 1;
                    if (iCounter % 20 == 0)
                    {
                        break;
                    }
                } while (Win.GetHandle("3615 - OEPAD Warning") == IntPtr.Zero);
                iCounter = 0;

                DialogHandler("3615 - OEPAD Warning", "OK");

                while (Win.GetHandle("3615 - Address Entry") == IntPtr.Zero)
                {
                    Thread.Sleep(500);
                    iCounter += 1;
                    if (iCounter % 20 == 0)
                    {
                        Console.WriteLine("Address Entry not found..");
                        break;
                    }
                }
                iCounter = 0;

                DialogHandler("3615 - Address Entry", "&Save");

                if (Win.IsChecked(childhWnd[(int)OP.SO]) == true)
                {
                    EMail email = new EMail();
                    email.Send(new string[] { "treische@wesco.com" }, "C# email test", "Testing OSR program");

                }
                //lastSalesOrder = salesOrder;

                //do
                //{
                //    MH.sendKey(childhWnd[(int)OSR.OrderGrid], Keys.Down, true);
                //    Thread.Sleep(500);
                //    salesOrder.ContractNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.ContractNo]);
                //    salesOrder.CreatedBy = MH.GetWindowTextRaw(childhWnd[(int)OSR.CreatedBy]);
                //    salesOrder.Customer = MH.GetWindowTextRaw(childhWnd[(int)OSR.Customer]);
                //    salesOrder.CustomerPO = MH.GetWindowTextRaw(childhWnd[(int)OSR.CustPO]);
                //    salesOrder.CustPartNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.CustPartNum]);
                //    salesOrder.OrderNumber = MH.GetWindowTextRaw(childhWnd[(int)OSR.OrderNum]);
                //    salesOrder.OrderStatus = MH.GetWindowTextRaw(childhWnd[(int)OSR.OrderStatus]);

                //} while (lastSalesOrder.OrderNumber == salesOrder.OrderNumber);
            }
            else
            {
                return false;
            }


            return true;
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

    } //automation class

    public struct SalesOrder
    {
        public string Customer { get; set; }
        public string OrderNumber { get; set; }
        public string OrderStatus { get; set; }
        public string CustomerPO { get; set; }
        public string CreatedBy { get; set; }
        public string CustPartNumber { get; set; }
        public string ContractNumber { get; set; }
    }
}
