using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;

namespace CrmOutlookAddIn2013
{
    public partial class ShowCustomersRibbon
    {
        private void ShowCustomersRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ShowCustomers_Click(object sender, RibbonControlEventArgs e)
        {
            AppointmentItem item = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem as AppointmentItem;
            if (item != null)
            {
                ShowCustomersForm frmShowCustomers = new ShowCustomersForm(item);
                frmShowCustomers.Show();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("A problem occured when tried to find reference to Appointment Item. Please contact IT support.");
            }

        }
    }
}
