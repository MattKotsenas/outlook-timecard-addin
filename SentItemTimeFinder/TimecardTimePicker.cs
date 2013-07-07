using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace SentItemTimeFinder
{
    [ComVisible(true)]
    public class TimecardTimePicker : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public TimecardTimePicker()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SentItemTimeFinder.TimecardTimePicker.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void DateButton_Click(Office.IRibbonControl control)
        {
            var selector = Globals.ThisAddIn.CustomTaskPanes.Add(new DateSelector(), "Select Dates");
            selector.Visible = true;
        }

        #endregion

        #region Helpers
        public Bitmap GetPunchCardImage(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.PunchClock);
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
