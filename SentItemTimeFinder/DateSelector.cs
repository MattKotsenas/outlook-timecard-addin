using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SentItemTimeFinder
{
    public partial class DateSelector : UserControl
    {
        public DateSelector()
        {
            InitializeComponent();

            // Auto-select the last two weeks
            DateTime endDate = DateTime.Today;
            DateTime startDate = endDate.AddDays(-13);

            datePicker.SelectionStart = startDate;
            datePicker.SelectionEnd = endDate;
        }

        private static IEnumerable<DateTime> CreateDateRange(DateTime start, DateTime end)
        {
            DateTime date = start;
            while (date <= end)
            {
                yield return date;
                date = date.AddDays(1);
            }
        }

        private static string FormatTimes(DateTime date, DateTime? firstTime, DateTime? lastTime)
        {
            var d = date.ToShortDateString();
            var f = (firstTime.HasValue) ? firstTime.Value.ToShortTimeString() : "(no mail)";
            var l = (lastTime.HasValue) ? lastTime.Value.ToShortTimeString() : "(no mail)";

            return d + "  " + f + "  " + l;
        }

        private void Calculate_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear error messages
                errorLabel.Text = "";
                errorLabel.Visible = false;

                var startDate = datePicker.SelectionStart;
                var endDate = datePicker.SelectionEnd;

                dataLabel.Text = "";

                foreach (var date in CreateDateRange(startDate, endDate))
                {
                    var firstTime = Globals.ThisAddIn.GetDate(date, SortOrder.First);
                    var lastTime = Globals.ThisAddIn.GetDate(date, SortOrder.Last);

                    var formattedString = FormatTimes(date, firstTime, lastTime);

                    dataLabel.Text += formattedString;
                    dataLabel.Text += "\n";
                }
            }
            catch (Exception ex)
            {
                errorLabel.Text = ex.Message;
                errorLabel.Visible = true;
            }
        }
    }
}
