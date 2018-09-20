using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace CrmOutlookAddIn2013
{
    public partial class CrmSyncRibbonDesign
    {
        
        private void CrmSyncRibbonDesign_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDoSync_Click(object sender, RibbonControlEventArgs e)
        {
            OutlookCrmSyncForm syncFrm = new OutlookCrmSyncForm();
            syncFrm.Show();
        }
    }
}
