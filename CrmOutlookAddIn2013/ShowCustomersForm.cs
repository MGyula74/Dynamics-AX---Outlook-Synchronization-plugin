using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Principal;
using Microsoft.Office.Interop.Outlook;
using CrmOutlookAddIn2013.Model;
using System.Reflection;

namespace CrmOutlookAddIn2013
{
    public partial class ShowCustomersForm : Form
    {
        //WebOutlookCrm.OutlookService service = new WebOutlookCrm.OutlookService();
        private CrmWebServiceProxy _serviceProxy = null;
        private UserData userData;

        CrmWebServiceProxy serviceProxy
        {
            get
            {
                if (_serviceProxy != null)
                    return _serviceProxy;

                tbInfo.Text = "Connecting to WebCRM. Please wait...";
                tbInfo.Refresh();

                WindowsIdentity wi = WindowsIdentity.GetCurrent();

                string domainUserName = wi.Name;
                string networkDomain = domainUserName.Split('\\')[0] + ".Local";
                string networkAlias = domainUserName.Split('\\')[1];

                _serviceProxy = new CrmWebServiceProxy(networkDomain, networkAlias);

                if (_serviceProxy.IsTestMode())
                {
                    networkAlias = serviceProxy.GetTestNetworkAlias();
                }

                tbInfo.Text = "Getting user data...";
                tbInfo.Refresh();

                userData = _serviceProxy.GetUserData(networkDomain, networkAlias);
                if (userData == null)
                {
                    tbInfo.Text = "User not configured for CRM in AX!";
                    tbInfo.Refresh();

                    _serviceProxy.WriteInfo(String.Format("BEGIN: User not found: {0}, {1}", networkDomain, networkAlias));
                }

                return _serviceProxy;
            }
        }

        AppointmentItem item;
        bool isSalesDistrictListPopulated;

        public ShowCustomersForm()
        {
            InitializeComponent();
            this.Text = String.Format("Business Relations v{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
        }

        public ShowCustomersForm(AppointmentItem _item)
        {
            InitializeComponent();
            this.item = _item;
            InitializeDataGridView();
            InitializeContactsGridView();
            this.AcceptButton = button3;
            this.Text = String.Format("Business Relations v{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            SearchBusinessRelations();
        }

        private void SearchBusinessRelations()
        {
            String salesDistrictSelected = Convert.ToString(ddlSalesDistrict.SelectedValue);
            if (String.IsNullOrEmpty(salesDistrictSelected))
            {
                tbInfo.Text = "You must select a sales district first!";
                tbInfo.Refresh();

                return;
            }

            int numOfRecords = PopulateDataGridView();
            tbInfo.Text = String.Format("Downloaded {0} records.", numOfRecords);
        }

        private int PopulateDataGridView()
        {
            String salesDistrictSelected = Convert.ToString(ddlSalesDistrict.SelectedValue);
            if (String.IsNullOrEmpty(salesDistrictSelected))
            {
                tbInfo.Text = "You must select a sales district first!";
                tbInfo.Refresh();

                return 0;
            }

            tbInfo.Text = "Fetching customer data. Please wait...";
            tbInfo.Refresh();

            //WebOutlookCrm.Northwind.SMMBUSRELTABLE_DisplayDataTable tab =
            //    service.GetBusinessRelationsBySalesDistrict("BCE\\" + networkAlias, company, salesDistrictSelected, tbCustNameFilter.Text);
            List<BusRelData> tab = serviceProxy.GetBusinessRelationsBySalesDistrict("BCE\\" + userData.networkAlias, userData.COMPANY, 
                salesDistrictSelected, tbCustNameFilter.Text);

            dataGridView1.DataSource = tab;
            serviceProxy.WriteInfo("END: PopulateDataGridView. Number of busrels = " + tab.Count);

            return tab.Count;
        }

        private int PopulateContactsGridView()
        {
            tbInfo.Text = "Fetching contact data. Please wait...";
            tbInfo.Refresh();

            string partyIdTxt = getPartyIdSelected();
            if (String.IsNullOrEmpty(partyIdTxt)) return 0;

            //WebOutlookCrm.Northwind.CONTACTPERSONShortDataTable contactTable = serviceProxy.GetContactsByPartyID(partyIdTxt, company);
            List<ContactPerson> contactTable = serviceProxy.GetContactsByPartyID(partyIdTxt, userData.COMPANY);
            gridViewContacts.DataSource = contactTable;
            serviceProxy.WriteInfo("END: PopulateContactsGridView. Number of contacts = " + contactTable.Count);

            tbInfo.Text = String.Format("Number of contacts = {0} (partyId: {1})", contactTable.Count, partyIdTxt);
            tbInfo.Refresh();
            return contactTable.Count;
        }

        private int PopulateSalesDistrictList()
        {
            tbInfo.Text = "Fetching sales districts. Please wait...";
            tbInfo.Refresh();

            //WebOutlookCrm.SMMTABLES.SalesDistrictDataTable sdTable =
            //    service.GetUserSalesDistricts("BCE\\" + networkAlias, company);
            List<SalesDistrict> sdTable = serviceProxy.GetUserSalesDistricts("BCE\\" + userData.networkAlias, userData.COMPANY);

            ddlSalesDistrict.DataSource = sdTable;
            ddlSalesDistrict.DisplayMember = "DESCRIPTION";
            ddlSalesDistrict.ValueMember = "SALESDISTRICTID";

            tbInfo.Text = string.Format("{0} sales district(s) found for user '{1}'. ", sdTable.Count, userData.networkAlias);
            if (sdTable.Count == 0)
                tbInfo.Text += string.Format("Hint: See the 'Sales data records filter' in AX.");

            return sdTable.Count;
        }

        private int PopulateContactList()
        {
            string partyIdTxt = getPartyIdSelected();
            if (String.IsNullOrEmpty(partyIdTxt)) return 0;

            tbInfo.Text = "Fetching contact data. Please wait...";
            tbInfo.Refresh();

            //WebOutlookCrm.Northwind.CONTACTPERSONShortDataTable cpTable =
            //    service.GetContactsByPartyID(partyIdTxt, company);
            List<ContactPerson> cpTable = serviceProxy.GetContactsByPartyID(partyIdTxt, userData.COMPANY);

            return cpTable.Count;
        }

        private void InitializeDataGridView()
        {
            dataGridView1.AutoGenerateColumns = false;
            //dataGridView1.AutoSize = true;

            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridView1.BorderStyle = BorderStyle.Fixed3D;
            dataGridView1.ReadOnly = true;

            // Initialize and add a text box column.
            DataGridViewColumn column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "FULLNAME";
            column.HeaderText = "Business relation name";
            column.Name = "NAME";
            column.Width = 250;
            dataGridView1.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "BUSRELACCOUNT";
            column.HeaderText = "Bus Rel Account";
            column.Width = 125;
            dataGridView1.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "CUSTACCOUNT";
            column.HeaderText = "Customer Account";
            column.Width = 125;
            dataGridView1.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "PARTYID";
            column.Name = "PARTYID";
            column.HeaderText = "Identifier";
            column.Width = 100;
            dataGridView1.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "RELATIONTYPE";
            column.HeaderText = "Relation type";
            column.Name = "RELATIONTYPE";
            column.Width = 100;
            //column.Visible = false;
            dataGridView1.Columns.Add(column);

            dataGridView1.Width = 610;
            dataGridView1.Refresh();
            //dataGridView1.Rows.Add();
        }

        private void InitializeContactsGridView()
        {
            gridViewContacts.AutoGenerateColumns = false;
            //gridViewContacts.AutoSize = true;

            gridViewContacts.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            gridViewContacts.BorderStyle = BorderStyle.Fixed3D;
            gridViewContacts.ReadOnly = true;

            // Initialize and add a text box column.
            DataGridViewColumn column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "NAME";
            column.HeaderText = "Contact name";
            column.Name = "NAME";
            column.Width = 200;
            gridViewContacts.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "PHONE";
            column.HeaderText = "Phone";
            column.Name = "Phone";
            column.Width = 150;
            gridViewContacts.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "EMAIL";
            column.HeaderText = "Email";
            column.Name = "Email";
            column.Width = 150;
            gridViewContacts.Columns.Add(column);

            column = new DataGridViewTextBoxColumn();
            column.DataPropertyName = "ContactPersonId";
            column.HeaderText = "Contact ID";
            column.Width = 100;
            column.Visible = false;     
            gridViewContacts.Columns.Add(column);
            gridViewContacts.Width = 610;
            gridViewContacts.Refresh();
            //gridViewContacts.Rows.Add();
        }

        private string getPartyIdSelected()
        {
            string partyIdTxt = String.Empty;
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView1.SelectedRows[0];

                partyIdTxt = Convert.ToString(row.Cells["PARTYID"].Value);
            }
            return partyIdTxt;
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            string brName = string.Empty;
            string partyId = string.Empty;
            string contactPersonId = string.Empty;
            string contactName = string.Empty;
            string contactPhone = string.Empty;
            string contactEmail = string.Empty;

            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView1.SelectedRows[0];
                brName = Convert.ToString(row.Cells["NAME"].Value);
                partyId = Convert.ToString(row.Cells["PARTYID"].Value);
            }

            if (gridViewContacts.SelectedRows.Count > 0)
            {
                DataGridViewRow row = (DataGridViewRow)gridViewContacts.SelectedRows[0];
                contactPersonId = Convert.ToString(row.Cells[3].Value);
                contactName = Convert.ToString(row.Cells[0].Value);
                contactPhone = Convert.ToString(row.Cells[1].Value);
                contactEmail = Convert.ToString(row.Cells[2].Value);
            }

            string sdValue = Convert.ToString(ddlSalesDistrict.SelectedValue);
            string sdText = ddlSalesDistrict.Text;

            /*item.Body += System.Environment.NewLine;
            item.Body += String.Format("## WebCRM ==> ## ");
            item.Body += String.Format("Business Relation Name: {0}; ", brName);
            item.Body += String.Format("Identifier: {0}; ", partyId);
            item.Body += String.Format("Contact: {0}; ", contactPersonId);
            item.Body += String.Format("## <== WebCRM ##");
            item.Body += System.Environment.NewLine;*/

            UserProperties ups = item.UserProperties;
            UserProperty prop = ups.Add("PartyId", OlUserPropertyType.olText, false, null);
            prop.Value = partyId;
            prop = ups.Add("ContactPersonId", OlUserPropertyType.olText, false, null);
            prop.Value = contactPersonId;

            //------------------------------------------------------
            string bodyText = item.Body;
            if (bodyText == null) bodyText = String.Empty;

            //String custDataBegin = "##<--Customer Data--##";
            String custDataBegin = "## CRM Data ==> ##";
            //String custDataEnd = "##--Customer Data-->##";
            String custDataEnd = "## <== CRM Data ##";
            int cdStartPos = bodyText.IndexOf(custDataBegin);
            int cdEndPos = bodyText.IndexOf(custDataEnd, cdStartPos == -1 ? 0 : cdStartPos);
            if (cdStartPos != -1 && cdEndPos != -1)
            {
                cdEndPos += custDataEnd.Length;
                bodyText = bodyText.Remove(cdStartPos, cdEndPos - cdStartPos);
            }
            if (cdStartPos == -1) cdStartPos = bodyText.Length;
            String custDataTxt = String.Empty;

            String custDataInfo = String.Empty;
            custDataInfo += String.Format("Customer name: {0}", brName) + Environment.NewLine;
            custDataInfo += String.Format("Identifier: {0}", partyId) + Environment.NewLine;
            custDataInfo += String.Format("Customer ABC: {0}", "-") + Environment.NewLine;
            custDataInfo += String.Format("Sales Rep: {0}", String.Format("{0} ({1})", sdText, sdValue)) + Environment.NewLine;
            custDataInfo += String.Format("Mode of delivery: {0}", "-") + Environment.NewLine;
            custDataInfo += String.Format("Search name: {0}", "-") + Environment.NewLine;
            custDataInfo += String.Format("Contact ID: {0}", contactPersonId) + Environment.NewLine;
            custDataInfo += String.Format("Contact name: {0}", contactName) + Environment.NewLine;
            custDataInfo += String.Format("Contact phone : {0}", contactPhone) + Environment.NewLine;
            custDataInfo += String.Format("Contact email : {0}", contactEmail);

            custDataTxt = Environment.NewLine;
            custDataTxt += custDataBegin + Environment.NewLine;
            custDataTxt += custDataInfo + Environment.NewLine;
            custDataTxt += custDataEnd;

            bodyText = bodyText.Insert(cdStartPos, custDataTxt);
            item.Body = bodyText;
            //------------------------------------------------------            

            this.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            String title = this.Text;
            if (this.Text.Contains(':'))
            {
                title = this.Text.Split(':')[0];
            }
            if (dataGridView1.SelectedRows.Count > 0)
            {
                this.Text = String.Format("{0}: {1}", title, getPartyIdSelected());
                btnSelect.Enabled = true;
            }
            else
                btnSelect.Enabled = false;

            if (dataGridView1.SelectedRows.Count > 0)
            {
                string partyIdTxt = getPartyIdSelected();
                if (String.IsNullOrEmpty(partyIdTxt))
                    btnSelect.Enabled = false;
                else
                {
                    PopulateContactsGridView();
                }
            }
        }

        private void tbCustNameFilter_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchBusinessRelations();
                e.Handled = true;
            }
        }

        private void ddlSalesDistrict_Click(object sender, EventArgs e)
        {
            if (!isSalesDistrictListPopulated)
            {
                PopulateSalesDistrictList();
                isSalesDistrictListPopulated = true;
            }
        }

        private void ddlSalesDistrict_SelectedIndexChanged(object sender, EventArgs e)
        {
            SalesDistrict row = (SalesDistrict)ddlSalesDistrict.SelectedItem;
            //WebOutlookCrm.SMMTABLES.SalesDistrictRow row = (WebOutlookCrm.SMMTABLES.SalesDistrictRow)drv.Row;
            if (!String.IsNullOrEmpty(row.SALESDISTRICTID))
            {
                tbInfo.Text = String.Format("Sales district: {0} ({1})", row.DESCRIPTION, row.SALESDISTRICTID);
                tbInfo.Refresh();

                dataGridView1.DataSource = null;
                gridViewContacts.DataSource = null;
            }

        }

        private void tbCustNameFilter_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
