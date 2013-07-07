using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace SentItemTimeFinder
{
    public enum SortOrder
    {
        First,
        Last
    }
    // TODO: Test on Outlook 2010 and with correct version of .NET
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public DateTime? GetDate(DateTime date, SortOrder order)
        {
            string startDate = date.ToShortDateString();
            string endDate = date.Date.AddDays(1).ToShortDateString();

            Outlook.MAPIFolder sentMail = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
            Outlook.Items items = sentMail.Items;
            Outlook.Items searchItems = items.Restrict("[SentOn] >= '" + startDate + "' AND [SentOn] <= '" + endDate + "'");

            var mails = new List<dynamic>();

            foreach (dynamic mailItem in searchItems)
            {
                if (mailItem != null)
                {
                    mails.Add(mailItem);
                }
            }

            DateTime? retVal;

            mails = mails.OrderBy(mail => mail.SentOn).ToList();

            if (mails.Count == 0)
            {
                retVal = null;
            }
            else
            {
                retVal = (order == SortOrder.First) ? mails.First().SentOn : mails.Last().SentOn;
            }

            return retVal;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new TimecardTimePicker();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
