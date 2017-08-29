using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using PInvoke;
using System.ComponentModel;


namespace ZDAT.Automation
{
    class ConvertOrder : Task
    {
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
                } while (MH.GetWindowTextRaw(childhWnd[OP.Customer]) == "");

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

                if (Win.IsChecked(childhWnd[OP.SO]) == true)
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
}
