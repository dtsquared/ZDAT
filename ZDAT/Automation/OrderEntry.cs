using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using PInvoke;
using System.ComponentModel;


namespace ZDAT.Automation
{
    public class OrderEntry : Task
    {
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
                            Mouse.LeftClick(OPchildhWnd[OP.OrderHeader]);

                            do
                            {
                                if (iCounter % 25 == 0)
                                {
                                    Mouse.LeftClick(OPchildhWnd[OP.OrderHeader]);
                                }
                                Thread.Sleep(300);
                                iCounter += 1;
                            } while (Win.IsOverlapped(OHhandle) == true);

                            iCounter = 1;
                            MH.setFocus(OHchildhWnd[OH.CustomerPO]);
                            Thread.Sleep(500);
                            MH.sendString(OHchildhWnd[OH.CustomerPO], lastPO);
                            Thread.Sleep(500);
                            if (EDIEnabled == true)
                            {
                                MH.setFocus(OHchildhWnd[OH.EdiPO]);
                                Thread.Sleep(250);
                                MH.sendString(OHchildhWnd[OH.EdiPO], lastPO);
                                Thread.Sleep(250);
                                MH.sendKey(OHchildhWnd[OH.EdiPO], Keys.Tab, true);
                                Thread.Sleep(300);
                            }

                            Mouse.LeftClick(OHchildhWnd[OH.Order]);
                            Thread.Sleep(300);
                            Mouse.LeftClick(OHchildhWnd[OH.Order]);
                            #endregion Fill_Order_Header

                            #region Save_Order
                            Console.Write("Saving Order");
                            Phandle = IntPtr.Zero;
                            Mouse.LeftClick(OHchildhWnd[OH.Save]);
                            do
                            {
                                if (iCounter % 25 == 0)
                                {
                                    Mouse.LeftClick(OHchildhWnd[OH.Save]);
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

                            if (Phandle != IntPtr.Zero)
                            {
                                PchildhWnd = Win.GetChildWindows(Phandle);
                                for (int i = 0; i < PchildhWnd.Count; i++)
                                {
                                    if (MH.GetWindowTextRaw(PchildhWnd[i]) == "Se&lect")
                                    {
                                        Thread.Sleep(500);
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
                                        break;
                                    }
                                }
                                Console.WriteLine();
                                Phandle = IntPtr.Zero;
                                iCounter = 0;

                                Console.Write("Confirming ticket printed");
                                for (int i = 0; i <= 25; i++)
                                {
                                    Phandle = Win.GetHandle("O111C");
                                    Thread.Sleep(250);

                                    if (Phandle != IntPtr.Zero)
                                        break;
                                }

                                if (Phandle != IntPtr.Zero)
                                {
                                    MH.setFocus(Phandle);
                                    PchildhWnd = Win.GetChildWindows(Phandle);
                                    for (int i = 0; i < PchildhWnd.Count; i++)
                                    {
                                        if (MH.GetWindowTextRaw(PchildhWnd[i]) == "OK")
                                        {
                                            Mouse.LeftClick(PchildhWnd[i]);
                                            Thread.Sleep(500);
                                            break;
                                        }
                                    }
                                }

                                for (int i = 0; i <= 25; i++)
                                {
                                    Thread.Sleep(250);
                                    if (MH.GetWindowTextRaw(OHchildhWnd[OH.CustomerPO]) == "")
                                        break;
                                    Console.Write(i + ", ");
                                }

                                Console.WriteLine();
                                Phandle = IntPtr.Zero;
                                iCounter = 1;
                            }
                            #endregion
                            #endregion

                            Mouse.LeftClick(OHchildhWnd[OH.OrderPad]);
                            do
                            {
                                if (iCounter % 20 == 0)
                                {
                                    Win.BringToTop(OPhandle);
                                    Mouse.LeftClick(OHchildhWnd[OH.OrderPad]);
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
                            if (MH.GetWindowTextRaw(OPchildhWnd[OP.Customer]) == "")
                            {
                                do
                                {
                                    MH.sendString(OPchildhWnd[OP.Customer], o.Customer);
                                    MH.sendKey(OPchildhWnd[OP.Customer], Keys.Enter, true);
                                    lastPO = o.PONumber;
                                    newOrder = true;
                                }
                                while (MH.GetWindowTextRaw(OPchildhWnd[OP.Customer]) != o.Customer);

                                do
                                {
                                    if (iCounter % 10 == 0)
                                    {
                                        MH.setFocus(OPchildhWnd[OP.Customer]);
                                        MH.sendKey(OPchildhWnd[OP.Customer], Keys.Enter, true);
                                    }
                                    Thread.Sleep(250);
                                    iCounter += 1;
                                } while (MH.GetWindowTextRaw(OPchildhWnd[OP.ShipToDesc]) == "");

                                if (MH.GetWindowTextRaw(OPchildhWnd[OP.LineIncr]) != "2")
                                {
                                    MH.sendString(OPchildhWnd[OP.LineIncr], "2");
                                }
                                iCounter = 1;
                            }

                            Thread.Sleep(250);

                            #endregion

                            #region Ship_To
                            if (o.Area.ToString() != MH.GetWindowTextRaw(OPchildhWnd[OP.ShipToNum]).ToString())
                            {
                                MH.setFocus(OPchildhWnd[OP.ShipToNum]);
                                MH.sendString(OPchildhWnd[OP.ShipToNum], o.Area);
                                MH.setFocus(OPchildhWnd[OP.ShipToNum]);
                                MH.sendKey(OPchildhWnd[OP.ShipToNum], Keys.Enter, true);
                                Thread.Sleep(250);

                                do
                                {
                                    if (iCounter % 25 == 0)
                                    {
                                        MH.sendKey(OPchildhWnd[OP.ShipToNum], Keys.Enter, true);
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
                            MH.setFocus(OPchildhWnd[OP.Qty]);
                            MH.sendString(OPchildhWnd[OP.Qty], o.Qty);
                            MH.sendKey(OPchildhWnd[OP.Qty], Keys.Enter, true);
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
                                    MH.setFocus(OPchildhWnd[OP.Product]);
                                    Thread.Sleep(300);
                                    MH.sendString(OPchildhWnd[OP.Product], o.Part.PadLeft(11, '0'));
                                    Thread.Sleep(250);
                                    MH.sendKey(OPchildhWnd[OP.Product], Keys.Enter, true);
                                    newOrder = false;
                                }
                                else
                                {
                                    if (MH.GetWindowTextRaw(OPchildhWnd[OP.ProdCode]) == "")
                                    {
                                        MH.sendString(OPchildhWnd[OP.Product], o.Part);
                                    }
                                    MH.sendString(OPchildhWnd[OP.Product], o.Part);
                                    MH.sendKey(OPchildhWnd[OP.Product], Keys.Tab, true);
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
                                        MH.setFocus(OPchildhWnd[OP.Product]);
                                        Thread.Sleep(300);
                                        if (MH.GetWindowTextRaw(OPchildhWnd[OP.ProdCode]) == "")
                                        {
                                            MH.sendString(OPchildhWnd[OP.Product], o.Part);
                                        }
                                        MH.sendKey(OPchildhWnd[OP.Product], Keys.Enter, true);
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
                                } while (MH.GetWindowTextRaw(OPchildhWnd[OP.Desc]) == "");
                                Phandle = IntPtr.Zero;
                                iCounter = 1;
                                Console.WriteLine();
                                #endregion

                                #region Price
                                //Mouse.LeftClick(childhWnd[OP.Price]);

                                #region Price Overrides
                                ////////////////////
                                //                                                                          //
                                // Used for price overrides - Tested ~140 line items with 100% success rate //
                                //                                                                          //
                                ////////////////////

                                //do
                                //{
                                //    //Console.WriteLine("Waiting for price to be set..");
                                //    MH.sendWMString(childhWnd[OP.Price], MH.WM_SETTEXT, "0.001");
                                //    Double.TryParse(MH.GetWindowTextRaw(childhWnd[OP.Price]), out Price);
                                //    Thread.Sleep(250);
                                //} while (Price <= 0.000);
                                #endregion

                                //MH.sendKey(childhWnd[OP.Price], Keys.Enter, true);
                                //Thread.Sleep(500);
                                #endregion

                                // Give some time for loading
                                Thread.Sleep(500);

                                // Add Date
                                Console.WriteLine("Setting Customer Delivery Date");
                                IntPtr CustDelDtHandle = OPchildhWnd[OP.CustDelDt];
                                bool dateError = true;

                                for (int i = 1; i <= 30; i++)
                                {
                                    if (dateError)
                                    {
                                        MH.setFocus(CustDelDtHandle);
                                        Thread.Sleep(250);
                                        MH.sendString(CustDelDtHandle, string.Format("{0:MM/dd/yy}", DateTime.Today.AddDays(i)));
                                        MH.sendKey(CustDelDtHandle, Keys.Tab, false);
                                    }
                                    else
                                        break;

                                    // Wait just in case the warning takes a while to load
                                    Thread.Sleep(500);

                                    Phandle = Win.GetHandle("CTLSHPVIAMT Warning");
                                    if (Phandle != IntPtr.Zero)
                                    {
                                        PchildhWnd = Win.GetChildWindows(Phandle);
                                        for (int j = 0; j < PchildhWnd.Count; j++)
                                        {
                                            if (MH.GetWindowTextRaw(PchildhWnd[j]) == "OK")
                                            {
                                                MH.setFocus(Phandle);
                                                Mouse.LeftClick(PchildhWnd[j]);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                        dateError = false;
                                }

                                #region Add
                                Console.WriteLine("Adding to order.");
                                Thread.Sleep(500);
                                Mouse.LeftClick(OPchildhWnd[OP.Add]);
                                Console.Write("Waiting for application to enter a ready state");
                                do
                                {
                                    Console.Write(".");

                                    if (MH.GetWindowTextRaw(OPchildhWnd[OP.Add]) == "Upd&ate")
                                    {
                                        Mouse.LeftClick(OPchildhWnd[OP.NewItem]);
                                        MissedLines.Add("LN: " + LineNumber + " Qty: " + o.Qty + " SIM: " + o.Part);
                                        Thread.Sleep(2000);
                                    }
                                    if (iCounter % 25 == 0)
                                    {
                                        Mouse.LeftClick(OPchildhWnd[OP.Add]);
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
                                } while (MH.GetWindowTextRaw(OPchildhWnd[OP.Qty]) != "0");

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
                    Mouse.LeftClick(OPchildhWnd[OP.OrderHeader]);

                    do
                    {
                        if (iCounter % 25 == 0)
                        {
                            Mouse.LeftClick(OPchildhWnd[OP.OrderHeader]);
                        }
                        Thread.Sleep(250);
                        iCounter += 1;
                    } while (Win.IsOverlapped(OHhandle) == true);
                    iCounter = 1;

                    MH.setFocus(OHchildhWnd[OH.CustomerPO]);
                    Thread.Sleep(500);
                    MH.sendString(OHchildhWnd[OH.CustomerPO], lastPO);
                    Thread.Sleep(500);
                    if (EDIEnabled == true)
                    {
                        MH.setFocus(OHchildhWnd[OH.EdiPO]);
                        Thread.Sleep(250);
                        MH.sendString(OHchildhWnd[OH.EdiPO], lastPO);
                        Thread.Sleep(250);
                        MH.sendKey(OHchildhWnd[OH.EdiPO], Keys.Tab, true);
                        Thread.Sleep(300);
                    }

                    Mouse.LeftClick(OHchildhWnd[OH.Order]);

                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        #region Save_Order
                        Console.Write("Saving Order");
                        Phandle = IntPtr.Zero;
                        Mouse.LeftClick(OHchildhWnd[OH.Save]);
                        do
                        {
                            if (iCounter % 25 == 0)
                            {
                                Mouse.LeftClick(OHchildhWnd[OH.Save]);
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
                            Mouse.LeftClick(OHchildhWnd[OH.OrderPad]);
                            do
                            {
                                if (iCounter % 15 == 0)
                                {
                                    Win.BringToTop(OHhandle);
                                    Thread.Sleep(250);
                                    Mouse.LeftClick(OPchildhWnd[OH.OrderPad]);
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
    }
}
