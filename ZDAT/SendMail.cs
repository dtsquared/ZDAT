using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ZDAT
{
    public class EMail
    {
        public void Send(string[] Recipients, string Subject, string Body)
        {
            try
            {
                Outlook.Application oApp = GetApplicationObject();

                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                Outlook.Recipients oRecips;

                oMsg.HTMLBody = Body;
                oMsg.Subject = Subject;

                foreach (string recip in Recipients)
                {
                    oMsg.Recipients.Add(recip);
                }

                oRecips = (Outlook.Recipients)oMsg.Recipients;
                oRecips.ResolveAll();
                ((Outlook._MailItem)oMsg).Send();
            }
            catch (Exception ex)
            {
                if (ex != null)
                {
                    System.Windows.Forms.MessageBox.Show("Unable to send mail.");
                }
            }
        }

        private Outlook.Application GetApplicationObject()
        {
            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                Console.WriteLine("Outlook found!");
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {
                Console.WriteLine("Starting outlook..");
                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("Outlook", "", Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }
    }
}
