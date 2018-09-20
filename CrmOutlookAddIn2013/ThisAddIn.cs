using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace CrmOutlookAddIn2013
{
    public partial class ThisAddIn
    {
        //private Office.CommandBar _objMenuBar;
        //private Office.CommandBarPopup _objNewMenuBar;
        //private Office.CommandBarButton _objButton;

        private string menuTag = "Outlook - CRM Synchronization";

        private bool syncActivitiesUsingForm = true;

        private string _id;
        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.MyMenuBar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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

        #region "Outlook2013 Menu"
        /*private void MyMenuBar()
        {
            this.EraseMyMenuBar();

            try
            {
                //Define the existent Menu Bar
                _objMenuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                string menuCaption = "&CRM";
                int controlCount = _objMenuBar.Controls.Count;

                //Define the new Menu Bar into the old menu bar
                _objNewMenuBar = (Office.CommandBarPopup)
                             _objMenuBar.Controls.Add(Office.MsoControlType.msoControlPopup
                                                    , missing
                                                    , missing
                                                    , missing
                                                    , true);

                if (_objNewMenuBar != null)
                {
                    _objNewMenuBar.Caption = menuCaption;
                    _objNewMenuBar.Tag = menuTag;

                    _objButton = (Office.CommandBarButton)_objNewMenuBar.Controls.
                        Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
                    _objButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    _objButton.Caption = "&Synchronize activities";
                    _objButton.Tag = "Button1";
                    _objButton.FaceId = 136;
                    _objButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(_objButton_Click);

                    _objNewMenuBar.Visible = true;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString(), "Error Message");
            }

        }*/

        #endregion

        #region "Event Handler"

        #region "Menu Button"

        private void _objButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            try
            {
                this.syncActivitiesWithCrm();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error " + ex.Message.ToString());
            }
        }

        #endregion

        #region "ToolBar Button"

        private void performSyncUsingForm()
        {
            //OutlookCrmSyncForm syncFrm = new OutlookCrmSyncForm();
            //syncFrm.Show();
        }

        private void syncActivitiesWithCrm()
        {
            try
            {
                if (syncActivitiesUsingForm)
                {
                    this.performSyncUsingForm();
                    return;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString());
            }
        }

        #endregion

        #region "Remove Existing"

        private void EraseMyMenuBar()
        {
            // If the menu already exists, remove it.
            try
            {
                Office.CommandBarPopup _objIsMenueExist = (Office.CommandBarPopup)
                    this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.
                    FindControl(Office.MsoControlType.msoControlPopup
                              , missing
                              , menuTag
                              , true
                              , true);

                if (_objIsMenueExist != null)
                {
                    _objIsMenueExist.Delete(true);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString()
                                                   , "Error Message");
            }
        }

        #endregion

        #endregion
    }
}
