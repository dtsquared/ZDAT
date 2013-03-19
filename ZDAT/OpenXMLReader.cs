using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ZDAT
{
    public class OpenXMLReader
    {
        public static IEnumerable<Order> ReadOXLFile(string filepath)
        {
            //Excel object references
            Workbook wb;
            IEnumerable<Sheet> workSheets;
            WorksheetPart orderSheet;
            SharedStringTable sharedStrings;

            //Declare helper variables.
            string wsID;
            List<Order> orders;

            //Open the excel workbook
            try
            {
                SpreadsheetDocument document = SpreadsheetDocument.Open(filepath, false);


                //References to the workbook and Shared String Table.
                wb = document.WorkbookPart.Workbook;
                workSheets = wb.Descendants<Sheet>();
                sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                //Reference to Excel worksheet with order data.
                wsID = workSheets.First(sheet => sheet.Name == @"Sheet1").Id;
                orderSheet = (WorksheetPart)document.WorkbookPart.GetPartById(wsID);

                //Load order data to business object.
                orders = Order.LoadOrders(orderSheet.Worksheet, sharedStrings);


                //LINQ Query for all orders
                IEnumerable<Order> allOrders = from order in orders select order;
                document.Close();
                return allOrders;
            }
            catch (Exception ex)
            {
                if (ex is FormatException)
                {
                    MessageBox.Show("Wrong format. Please make sure the file is saved in xlsx format.");
                }
                if (ex is IOException)
                {
                    MessageBox.Show("Could not open file. Please make sure it is not already in use.");
                }
                if (ex is System.InvalidOperationException)
                {
                    MessageBox.Show("Invalid format. Please make sure you are trying to load the correct file.");
                }
                return new List<Order>();
            }
        }

        public class Order
        {
            //Properties.
            public string PONumber { get; set; }
            public string Customer { get; set; }
            public string Area { get; set; }
            public string Part { get; set; }
            public string Qty { get; set; }

            /// <summary>
            /// Helper method for creating a list of orders 
            /// from an Excel worksheet.
            /// </summary>
            public static List<Order> LoadOrders(Worksheet worksheet, SharedStringTable sharedString)
            {
                //Initialize order list.
                List<Order> result = new List<Order>();

                //LINQ query to skip first row with column names.
                IEnumerable<Row> dataRows =
                  from row in worksheet.Descendants<Row>()
                  where row.RowIndex > 1
                  select row;

                foreach (Row row in dataRows)
                {
                    //LINQ query to return the row's cell values.
                    //Where clause filters out any cells that do not contain a value.
                    //Select returns cell's value unless the cell contains
                    //  a shared string.
                    //If the cell contains a shared string its value will be a 
                    //  reference id which will be used to look up the value in the 
                    //  shared string table.
                    IEnumerable<String> textValues =
                      from cell in row.Descendants<Cell>()
                      where cell.CellValue != null
                      select (cell.DataType != null
                      && cell.DataType.HasValue
                      && cell.DataType == CellValues.SharedString ?
                      sharedString.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText : cell.CellValue.InnerText);

                    //Check to verify the row contains data.
                    if (textValues.Count() > 0)
                    {
                        //Create an Order and add it to the list.
                        var textArray = textValues.ToArray();
                        Order order = new Order();

                        order.PONumber = textArray[0];
                        order.Part = textArray[1].PadLeft(11, '0');
                        order.Area = textArray[2];
                        order.Qty = textArray[3];
                        order.Customer = textArray[4];

                        result.Add(order);
                    }
                    else
                    {
                        //If no cells, then you have reached the end of the table.
                        break;
                    }
                }

                //Return populated list of orders.
                return result;
            }
        }

    }

    public class TextFileRW
    {
        public static string GetSWQ()
        {
            string SWQ;
            StreamReader sr = new StreamReader(@".\last.txt");
            SWQ = sr.ReadLine();
            sr.Close();
            return SWQ;
        }

        public static void WriteSWQ(string SWQ)
        {
            StreamWriter sr = new StreamWriter(@".\last.txt");
            sr.Write(SWQ);
            sr.Close();
        }
    }
}
