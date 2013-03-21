using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Diagnostics;
using System.Threading;
using System.ComponentModel;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Threading;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;

namespace ZDAT
{
    public partial class MainWindow : Window
    {
        #region Public_Variables
        //Message queues for retreiving info from background worker threads
        public static ConcurrentQueue<int> cqSOM = new ConcurrentQueue<int>();
        public static ConcurrentQueue<string> cqEDI = new ConcurrentQueue<string>();

        public string Branch
        {
            get
            {
                return branch;
            }
            set
            {
                if (value.Length < 4)
                    MessageBox.Show("Branch number must be 4 digits");
                else
                    branch = ZDAT.Properties.Settings.Default.Branch;
            }
        }
        #endregion

        #region Private_Variables
        //Background Workers
        private BackgroundWorker bgWorkerSOM = new BackgroundWorker();
        private BackgroundWorker bgWorkerEDI = new BackgroundWorker();

        private IEnumerable<OpenXMLReader.Order> Orders { get; set; }   //SOM order collection
        private List<OpenXMLReader.Order> orders { get; set; }  //SOM order converted to list

        private CustomerRange custRange = new CustomerRange();  //EDI customer dpc range
        private bool EDIEnabled = false;    //Passed to bgWorkerSOM for EDI enabled customers
        private int LineCount;  //Number of lines on SOM order
        private string branch; //Branch number
        #endregion

        public class CustomerRange
        {
            public string CustTo { get; set; }
            public string CustFrom { get; set; }
        }

        #region Main_Window
        //Main Window initialization
        public MainWindow()
        {
            InitializeComponent();

            #region bgWorker_SOM
            bgWorkerSOM.DoWork += new DoWorkEventHandler(bgWorkerSOM_DoWork);
            bgWorkerSOM.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorkerSOM_RunWorkerCompleted);
            bgWorkerSOM.ProgressChanged += new ProgressChangedEventHandler(bgWorkerSOM_ProgressChanged);
            bgWorkerSOM.WorkerReportsProgress = true;
            bgWorkerSOM.WorkerSupportsCancellation = true;
            #endregion

            #region bgWorker_EDI
            bgWorkerEDI.DoWork += new DoWorkEventHandler(bgWorkerEDI_DoWork);
            bgWorkerEDI.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorkerEDI_RunWorkerCompleted);
            bgWorkerEDI.ProgressChanged += new ProgressChangedEventHandler(bgWorkerEDI_ProgressChanged);
            bgWorkerEDI.WorkerReportsProgress = true;
            bgWorkerEDI.WorkerSupportsCancellation = true;
            #endregion

        }

        private void MainWindow1_Closing(object sender, CancelEventArgs e)
        {
            if (txtBranch.Text != "")
            {
                ZDAT.Properties.Settings.Default.Branch = txtBranch.Text;
                ZDAT.Properties.Settings.Default.Save();
            }
        }
        #endregion

        #region tabOrderEntry
        private void tabOrderEntry_Loaded(object sender, RoutedEventArgs e)
        {
            Branch = ZDAT.Properties.Settings.Default.Branch;
            txtBranch.Text = Branch;
        }

        private void tabOrderEntry_GotFocus(object sender, RoutedEventArgs e)
        {
            MainWindow1.Width = 630;
        }

        private void txtBranch_LostFocus(object sender, RoutedEventArgs e)
        {
            Branch = txtBranch.Text;
        }

        private void imgDrop_Drop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (s.Length > 1)
            {
                MessageBox.Show("Please only use one file at a time.");
            }
            else
            {
                Orders = OpenXMLReader.ReadOXLFile(@s[0]);
                orders = Orders.ToList<OpenXMLReader.Order>();
                if (orders.Count > 0 && txtBranch.Text != "")
                {
                    dgDisplay.ItemsSource = orders;
                    LineCount = orders.Count;
                    btnEnterOrder.IsEnabled = true;
                    lblStatus.Content = "Order successfully loaded";
                }
                else
                {
                    lblStatus.Content = "Unable to load file.";
                }
            }
        }

        private void btnEnterOrder_Click(object sender, RoutedEventArgs e)
        {
            if (orders != null)
            {
                if (bgWorkerSOM.IsBusy != true)
                {
                    lblConsole.Content = orders[0].PONumber;
                    lblStatus.Content = "Entering PO number:";
                    btnEnterOrder.IsEnabled = false;
                    btnCancel.IsEnabled = true;
                    bgWorkerSOM.RunWorkerAsync(Orders);
                }
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (bgWorkerSOM.IsBusy == true)
            {
                lblStatus.Content = "Canceling...";
                btnCancel.IsEnabled = false;
                bgWorkerSOM.CancelAsync();
            }
        }
        #endregion

        #region tabInqOrd
        private void tabInqOrd_GotFocus(object sender, RoutedEventArgs e)
        {
            MainWindow1.Width = 447;
        }

        private void chkEDI_Checked(object sender, RoutedEventArgs e)
        {
            if (chkEDI.IsChecked == true)
                EDIEnabled = true;
            else
                EDIEnabled = false;
        }

        private void txtCustTo_GotFocus(object sender, RoutedEventArgs e)
        {
            txtCustTo.SelectAll();
        }

        private void txtCustFrom_LostFocus(object sender, RoutedEventArgs e)
        {
            txtCustTo.Text = txtCustFrom.Text;
        }

        private void txtEmail_GotFocus(object sender, RoutedEventArgs e)
        {
            btnAddEmail.IsDefault = true;
        }
        #endregion

        #region tabOptions
        private void tabOptions_GotFocus(object sender, RoutedEventArgs e)
        {
            MainWindow1.Width = 460;
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.ShowDialog();
            txtAcuthinPath.Text = dlg.FileName;
        }

        #endregion

        #region bgWorker_EDI
        private void bgWorkerEDI_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void bgWorkerEDI_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((bool)e.Result == false)
            {
                MessageBox.Show("Window not found!\r\nPlease make sure order seach is open.", "Error");
            }
        }

        private void bgWorkerEDI_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Automation ae = new Automation();

            e.Result = ae.ConvertEDI(custRange, worker, e);
        }

        private void btnConvertEDI_Click(object sender, RoutedEventArgs e)
        {
            custRange.CustFrom = txtCustFrom.Text;
            custRange.CustTo = txtCustTo.Text;

            if (bgWorkerEDI.IsBusy != true)
            {
                bgWorkerEDI.RunWorkerAsync(custRange);
            }
        }
        #endregion

        #region bgWorker_SOM
        private void bgWorkerSOM_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Automation ae = new Automation();
             
            e.Result = ae.EnterOrders((IEnumerable<OpenXMLReader.Order>)e.Argument, LineCount, EDIEnabled, branch, worker, e);
        }

        private void bgWorkerSOM_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int result;

            ProgressBar1.Value = e.ProgressPercentage;

            cqSOM.TryDequeue(out result);
            lblConsole.Content = orders[result - 1].PONumber;
        }

        private void bgWorkerSOM_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                MainWindow1.Activate();
                MessageBox.Show("Canceled!");
                lblStatus.Content = "Stopped on PO:";
                btnEnterOrder.IsEnabled = true;
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error: " + e.Error.Message);
            }
            else
            {
                MainWindow1.Activate();
                MessageBox.Show("Done!");
                lblStatus.Content = "Last PO entered:";
                btnCancel.IsEnabled = false;
                btnEnterOrder.IsEnabled = true;
            }
        }
        #endregion


        private void btnMapScreen_Click(object sender, RoutedEventArgs e)
        {
            IntPtr hWnd = PInvoke.Win.GetHandle(txtWindowName.Text);
            List<IntPtr> childhWnd = PInvoke.Win.GetChildWindows(hWnd);

            for (int i = 0; i < childhWnd.Count; i++)
                PInvoke.MH.sendString(childhWnd[i], i.ToString());
        }



    }
}
