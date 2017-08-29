using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZDAT
{
    public struct OP
    {                                           // Field            Field Type      Area        Position
        public static int Customer = 116;       // Customer number  text field      top         left
        public static int OrderNum = 119;       // Order number     text field      top         left
        public static int Qty = 42;             // Quantity         text field      center      left
        public static int Product = 43;         // Product number   text field      center      top center
        public static int SO = 48;              // Item Type        radio button    center      center
        public static int ProdCode = 54;        // Product Code     text field      center      right
        public static int Desc = 130;           // Description      text Field      center      center
        public static int Add = 162;            // Add              button          center      bottom right
        public static int NewItem = 160;        // New Item         button          center      bottom right
        public static int Order = 156;          // Order Type       radio button    bottom      bottom center
        public static int OrderHeader = 124;    // Order Header     button          top         top center
        public static int ShipToNum = 101;      // Ship To Number   text field      top         top center
        public static int ShipToDesc = 99;      // Ship To Desc     text area       top         top center
        public static int ShipToOvr = 98;       // Ship To Ovr      button          top         top center
        public static int LineIncr = 18;        // Line Increment   text field      center      bottom
        public static int CustDelDt = 73;       // Cust Deliv Date  text field      center      bottom
    }

    public struct OH
    {
        public static int Customer = 116;
        public static int OrderNum = 118;
        public static int SoldTo = 135;
        public static int ShipToDesc = 138;
        public static int ShipToNum = 141;
        public static int OrderPad = 131;
        public static int CustomerPO = 41;
        public static int JobNumber = 42;
        public static int EdiPO = 51;
        public static int ShipTerms = 77;
        public static int ShipTermsDesc = 80;
        public static int ShipVia = 78;
        public static int ShipViaDesc = 81;
        public static int ShipAcctNum = 90;
        public static int ShipInstruc = 79;
        public static int Memo = 157;
        public static int Inquiry = 158;
        public static int Order = 159;
        public static int Save = 126;
        public static int Clear = 127;
        public static int Close = 134;
    }

    public struct OS
    {
        public static int ByItem = 2;
        public static int itmCatalogNum = 3;
        public static int itmCustPartNum = 4;
        public static int itmSIMNum = 6;
        public static int itmNoneOfAbove = 7;
        public static int ByOrder = 16;
        public static int ordCustPONum = 8;
        public static int ordOrderedBy = 9;
        public static int ordQuotedTo = 10;
        public static int ordJobNumber = 11;
        public static int ordInsideSales = 12;
        public static int ordOutsideSales = 13;
        public static int ordInvoiceNum = 14;
        public static int ordNoneOfAbove = 15;
        public static int fltrCustFrom = 18;
        public static int fltrCustTo = 21;
        public static int fltrShipFrom = 25;
        public static int fltrShipTo = 27;
        public static int fltrItemSchShipDateFrom = 31;
        public static int fltrItemSchShipDateTo = 33;
        public static int fltrItemReqDateFrom = 37;
        public static int fltrItemReqDateTo = 39;
        public static int fltrOrdDateFrom = 43;
        public static int fltrORdDateTo = 45;
        public static int ClearFilterBy = 47;
        public static int otypeAll = 50;
        public static int otypeNormal = 51;
        public static int otypeShipComplete = 52;
        public static int otypeAssembleHold = 53;
        public static int otypeShippedType = 54;
        public static int ordsAll = 55;
        public static int ordsOrders = 56;
        public static int ordsInquiries = 57;
        public static int ordsCreditMemo = 58;
        public static int ordsMemos = 59;
        public static int ordsOrdersInquiries = 60;
        public static int ordsOrdersCreditMemo = 61;
        public static int ostatusAll = 62;
        public static int ostatusOpen = 63;
        public static int ostatusClosed = 64;
        public static int shipstatAll = 65;
        public static int shipstatNotPicked = 66;
        public static int shipstatPicked = 67;
        public static int sipstatShipped = 68;
        public static int shipstatInvoiced = 69;
        public static int Search = 70;
        public static int Clear = 71;
        public static int CLose = 72;
    }

    public struct OSR
    {
        public static int OrderGrid = 2;
        public static int OrderNum = 13;
        public static int Cycle = 15;
        public static int OrderStatus = 17;
        public static int OrderType = 19;
        public static int OrderedBy = 25;
        public static int QuotedTo = 64;
        public static int ShipTerms = 32;
        public static int ShipTermsDesc = 34;
        public static int ShipVia = 33;
        public static int ShipViaDesc = 35;
        public static int InSales = 42;
        public static int OutSales = 43;
        public static int CustPO = 21;
        public static int JobNum = 23;
        public static int CreatedBy = 45;
        public static int InvoiceNum = 27;
        public static int Type = 18;
        public static int OrderTotal = 56;
        public static int Supplier = 60;
        public static int Customer = 6;
        public static int CustomerDesc = 7;
        public static int ViewRemoteOrder = 124;
        public static int ViewRemotePO = 125;
        public static int ViewReplCost = 126;
        public static int ProductNotes = 128;
        public static int ViewPO = 129;
        public static int ViewOrder = 130;
        public static int ManageOrder = 131;
        public static int ShippingActivity = 132;
        public static int NewSearch = 133;
        public static int Close = 134;
        public static int CustPartNum = 136;
        public static int SIMNum = 65;
        public static int ContractNo = 229;
        public static int CurPrice = 224;
        public static int LocalInventoryDesc = 110;
        public static int Summary = 135;
    }
}
