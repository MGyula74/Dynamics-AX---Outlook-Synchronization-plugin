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
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Reflection;
using Nager.Date;
using System.IO;
using System.Xml.Serialization;
using System.Xml.Linq;
using System.Xml;

namespace CrmOutlookAddIn2013
{
    public partial class OutlookCrmSyncForm : Form
    {
        //WebOutlookCrm.OutlookService service = new WebOutlookCrm.OutlookService();
        CrmWebServiceProxy serviceProxy;

        Outlook.Application outlookApp = new Outlook.Application();

        public int deletedCRMItemCounter = 0;
        public int updatedCRMItemCounter = 0;
        public int savedCRMItemCounter = 0;
        public int insertedOutlookItemCounter = 0;
        public int deletedOutlookItemCounter = 0;
        public int savedOutlookItemCounter = 0;

        public bool log { get; set; }

        String networkDomain;
        String networkAlias;
        public string outlookCalendarID { get; set; }
        public string outlookCalendarStoreID { get; set; }
        public string empId { get; private set; }
        public string company { get; set; }
        public int rangeDateBack { get; set; }
        public int rangeDateNext { get; set; }
        public DateTime lastSyncTime { get; set; }

        public string CompanyCountry { get; set; }

        Dictionary<string, string> smmActivityList;

        public OutlookCrmSyncForm()
        {
            InitializeComponent();
            this.Text = String.Format("Outlook - CRM Synchronization v{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            startSync();
        }

        private bool startSync()
        {
            Stopwatch swTotal = new Stopwatch();
            swTotal.Start();

            bool ret = true;

            deletedCRMItemCounter = 0;
            updatedCRMItemCounter = 0;
            savedCRMItemCounter = 0;

            insertedOutlookItemCounter = 0;
            deletedOutlookItemCounter = 0;
            savedOutlookItemCounter = 0;

            this.getNetworkCredentials();
            serviceProxy = new CrmWebServiceProxy(networkDomain, networkAlias);

            this.log = serviceProxy.IsLog();
            
            printLogMsg(string.Format("Connecting to {0} environment [{1}]", serviceProxy.isTest ? "Staging" : "Production", serviceProxy.Url.Replace("http://", String.Empty)));

            if (serviceProxy.IsTestMode())
            {
                networkAlias = serviceProxy.GetTestNetworkAlias();
            }

            Model.UserData userData = serviceProxy.GetUserData(networkDomain, networkAlias);
            
            if (userData == null)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("The user is not configured in Web CRM: {0}\\{1}", networkDomain, networkAlias));
                return false;
                //throw new Exception("The user data table is not set!");
            }

            if (log)
                serviceProxy.WriteInfo("++++++WebService: userDataTable.Count != 0");

            this.outlookCalendarID = userData.OUTLOOKCALENDAROUTLOOKENTRYID;
            this.outlookCalendarStoreID = string.IsNullOrEmpty(userData.OUTLOOKCALENDAROUTLOOKSTOREID) ? string.Empty : userData.OUTLOOKCALENDAROUTLOOKSTOREID;
            this.company = userData.COMPANY;
            this.empId = userData.EMPLID;
            //B19443_CrmChangesToPlugin MGY 2018.05.31 Begin
            this.CompanyCountry = userData.CompanyCountryId;
            //B19443_CrmChangesToPlugin MGY 2018.05.31 End

            //B12773_CRMOutlookPluginIssues MGY 2016.07.01 Begin
            //this.rangeDateBack = service.FilterDateBack();
            //this.rangeDateNext = service.FilterDateNext();
            this.rangeDateBack = userData.smmSynchronizeDaysBack;
            this.rangeDateNext = userData.smmSynchronizeDaysForward;
            if (this.rangeDateBack == 0) this.rangeDateBack = serviceProxy.FilterDateBack();
            if (this.rangeDateNext == 0) this.rangeDateNext = serviceProxy.FilterDateNext();
            this.lastSyncTime = serviceProxy.getLastSynchronizationTime(this.company, this.empId).ToLocalTime();
            //B12773_CRMOutlookPluginIssues MGY 2016.07.01 End

            if (string.IsNullOrEmpty(outlookCalendarID) || string.IsNullOrEmpty(outlookCalendarStoreID))
            {
                this.chooseCalendar();
            }

            printLogMsg(string.Format("Last synchronization time: {0}.", this.lastSyncTime));
            printLogMsg("Synchronization started.");

            if (!string.IsNullOrEmpty(this.company)
             && !string.IsNullOrEmpty(this.empId)
             && !string.IsNullOrEmpty(this.outlookCalendarID)
             && !string.IsNullOrEmpty(this.outlookCalendarStoreID))
            {
                String msg1, msg2 = String.Empty;

                Stopwatch sw1 = new Stopwatch();
                sw1.Start();

                msg1 = this.syncFromCrmToOutlook(company, empId);

                sw1.Stop();

                Stopwatch sw2 = new Stopwatch();
                sw2.Start();
                msg2 = this.syncFromOutlookToCrm();

                sw2.Stop();

                if (!String.IsNullOrEmpty(msg1) || !String.IsNullOrEmpty(msg2))
                    printLogMsg(String.Format("{0}\n{1}", msg1, msg2));

                //printLogMsg(String.Format("Crm -> OL: {0}", sw1.ElapsedMilliseconds));
                //printLogMsg(String.Format("OL -> Crm: {0}", sw2.ElapsedMilliseconds));
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                ret = false;
                //throw new Exception("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID, DoaminUser: " + networkAlias + "Domain: " + networkDomain + ")!");
            }

            //B12773_CRMOutlookPluginIssues MGY 2016.07.01 Begin
            serviceProxy.synchronizationFinished(this.company, this.empId, DateTime.Now);
            //B12773_CRMOutlookPluginIssues MGY 2016.07.01 End

            printLogMsg("Synchronization finished.");
            printLogMsg("--------------------------------------------------------");

            swTotal.Stop();
            //printLogMsg(String.Format("Total time: {0}", swTotal.ElapsedMilliseconds));

            return ret;
        }

        private String syncFromCrmToOutlook(string _company, string _responsibleEployee)
        {
            printLogMsg("Synchronizing activities from CRM to Outlook...");

            if (log)
                serviceProxy.WriteInfo("++++++ syncFromCrmToOutlook: CRM ---> Outlook");

            smmActivityList = new Dictionary<string, string>();
            bool isUpdateOutlookEntryID = false;
            int taskProirity = 0;
            int sensitivity = 0;
            int activityTimeType = 0;
            bool noSync = false;
            bool noAddList = false;

            try
            {
                //XElement booksFromFile = XElement.Load(@"c:\\act.xml");
                //String xmlStr = booksFromFile.ToString();
                String soapMsg = File.ReadAllText(@"c:\\act.xml");
                XmlDocument document = new XmlDocument();
                document.LoadXml(soapMsg);  //loading soap message as string
                XmlNamespaceManager manager = new XmlNamespaceManager(document.NameTable);

                manager.AddNamespace("d", "http://someURL");

                XmlNodeList xnList = document.SelectNodes("//bookHotelResponse", manager);
                int nodes = xnList.Count;

                //foreach (XmlNode xn in xnList)
                //{
                //    Status = xn["d:bookingStatus"].InnerText;
                //}

                XmlRootAttribute xRoot = new XmlRootAttribute();
                xRoot.ElementName = "DocumentElement";
                xRoot.IsNullable = true;
                WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable acts;
                XmlSerializer ser = new XmlSerializer(typeof(WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable), xRoot);
                StreamReader reader = new StreamReader("c:\\act.xml");
                acts = (WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable)ser.Deserialize(reader);
                reader.Close();

                Stopwatch swGetActs = new Stopwatch();
                swGetActs.Start();

                string Responsible = _responsibleEployee;
                /*if (Responsible.ToLower() == "krost")
                {
                    Responsible += "_test";
                }*/
                List<Model.ActivityData> smmAct = serviceProxy.GetsmmActivities(_company, Responsible);
                swGetActs.Stop();
                //printLogMsg(String.Format("service.GetsmmActivities: {0}", swGetActs.ElapsedMilliseconds));

                Outlook.AppointmentItem oAppointment = null;
                Outlook.Folder calFolder = GetFolder(outlookCalendarID, outlookCalendarStoreID);

                progressFromCrm.Minimum = 0;
                progressFromCrm.Maximum = 1;
                progressFromCrm.Value = 0;
                progressFromCrm.Step = 1;

                if (smmAct.Count == 0)
                {
                    progressFromCrm.PerformStep();
                }
                else
                {
                    progressFromCrm.Maximum = smmAct.Count;
                }

                foreach (Model.ActivityData item in smmAct)
                {
                    Stopwatch swAct = new Stopwatch();
                    swAct.Start();

                    if (item.STARTDATETIME > item.ENDDATETIME)
                    {
                        printLogMsg(string.Format("Wrong activity setting: end date is earlier than start date! act#: {0}, start: {1}, end: {2}", 
                            item.ACTIVITYNUMBER, item.STARTDATETIME.ToLocalTime(), item.ENDDATETIME.ToLocalTime()));
                        printLogMsg("Activity skipped.");
                        continue;
                    }

                    if (log)
                    {
                        //printLogMsg(String.Format("Checking activity: {0} (actnum={1})", item.PURPOSE, item.ACTIVITYNUMBER));
                        serviceProxy.WriteInfo("++++++ Activity found in CRM : ActivityNumber: " + item.ACTIVITYNUMBER);
                    }

                    noSync = false;
                    noAddList = false;
                    oAppointment = null;
                    bool insert = false;

                    Stopwatch swGetApp = new Stopwatch();
                    swGetApp.Start();
                    oAppointment = GetAppointment(item.OUTLOOKENTRYID, outlookCalendarStoreID);
                    swGetApp.Stop();

                    Stopwatch swCheckApp = new Stopwatch();
                    swCheckApp.Start();

                    if (oAppointment == null)
                    {
                        if (log)
                        {
                            serviceProxy.WriteInfo("++++++ Appointment not found in Outlook for Activity: " + item.ACTIVITYNUMBER);
                        }

                        if (string.IsNullOrEmpty(item.OUTLOOKENTRYID))
                        {
                            isUpdateOutlookEntryID = true;
                            insert = true;
                            oAppointment = (Outlook.AppointmentItem)calFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);

                            if (log)
                            {
                                serviceProxy.WriteInfo("++++++WebService: string.IsNullOrEmpty(item.OUTLOOKENTRYID)");
                                serviceProxy.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                                serviceProxy.WriteInfo("++++++WebService: noSync: " + noSync);
                            }
                        }
                        else
                        {
                            // ez az ág mikor nem találja az outlookban viszont van OutlookEntryID-ja tehát az outlookból lett törölve
                            if (log)
                            {
                                serviceProxy.WriteInfo("++++++WebService: string.IsNullOrEmpty(item.OUTLOOKENTRYID) != ");
                                serviceProxy.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                            }

                            noSync = true;
                            insert = false;
                            noAddList = serviceProxy.DeleteActivity(item.OUTLOOKENTRYID, company);

                            if (log)
                            {
                                printLogMsg("Activity removed from CRM: " + item.ACTIVITYNUMBER);
                                serviceProxy.WriteInfo("++++++WebService: " + "The " + item.ACTIVITYNUMBER + " activities deleted!");
                            }

                            deletedCRMItemCounter++;
                        }
                    }
                    //B19443_CrmChangesToPlugin MGY 2018.05.31 Begin
                    else if (oAppointment.Sensitivity != Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate)
                    //B19443_CrmChangesToPlugin MGY 2018.05.31 End
                    {
                        insert = false;

                        if (log)
                        {
                            serviceProxy.WriteInfo("++++++WebService: oAppointment != null");
                            serviceProxy.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                        }

                        isUpdateOutlookEntryID = false;

                        if (oAppointment.LastModificationTime > item.MODIFIEDDATETIME.ToLocalTime())
                        {
                            if (OutlookMgt.IsdifferentRecord(oAppointment, item))
                            {
                                if (log)
                                {
                                    serviceProxy.WriteInfo("++++++WebService: oAppointment.LastModificationTime > item.MODIFIEDDATETIME.ToLocalTime()");
                                    serviceProxy.WriteInfo("++++++WebService: " + oAppointment.LastModificationTime + " " + item.MODIFIEDDATETIME.ToLocalTime());
                                    serviceProxy.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                                }

                                taskProirity = 0;
                                sensitivity = 0;
                                activityTimeType = 0;

                                OutlookMgt.TASKPRIORITYConvertOutlookOlImportanceToInt(ref taskProirity, oAppointment);
                                OutlookMgt.SENSITIVITYConvertOutlookOlSensitivityToInt(ref sensitivity, oAppointment);
                                OutlookMgt.ACTIVITYTIMETYPEConvertOutlookOlBusyStatusToInt(ref activityTimeType, oAppointment);

                                noSync = serviceProxy.UpdateActivity(oAppointment.Start.ToUniversalTime(),
                                                                oAppointment.End.ToUniversalTime(),
                                                                empId,
                                                                oAppointment.AllDayEvent,
                                                                oAppointment.BillingInformation,
                                                                oAppointment.Body,
                                                                oAppointment.Subject,
                                                                oAppointment.Categories,
                                                                taskProirity,
                                                                oAppointment.Location,
                                                                oAppointment.Mileage,
                                                                oAppointment.ReminderSet,
                                                                oAppointment.ReminderMinutesBeforeStart,
                                                                oAppointment.Resources,
                                                                oAppointment.ResponseRequested,
                                                                sensitivity,
                                                                activityTimeType,
                                                                networkAlias,
                                                                company,
                                                                oAppointment.EntryID,
                                                                item.ACTIVITYNUMBER);
                                if (log)
                                {
                                    serviceProxy.WriteInfo("++++++WebService: " + "The " + item.ACTIVITYNUMBER + " activities updated!");
                                    printLogMsg("Activity modified in CRM: " + item.ACTIVITYNUMBER);
                                }
                                updatedCRMItemCounter++;
                            }
                        }
                    }
                    //B19443_CrmChangesToPlugin MGY 2018.05.31 Begin
                    else
                    {
                        noSync = true;
                        noAddList = serviceProxy.DeleteActivity(item.OUTLOOKENTRYID, company);

                        if (log)
                        {
                            printLogMsg("Appoointment has became Private, therefore Activity removed from CRM: " + item.ACTIVITYNUMBER);
                            serviceProxy.WriteInfo("++++++WebService: " + "The " + item.ACTIVITYNUMBER + " activities deleted!");
                        }

                        deletedCRMItemCounter++;
                    }
                    //B19443_CrmChangesToPlugin MGY 2018.05.31 End

                    swCheckApp.Stop();

                    Stopwatch swCheckApp2 = new Stopwatch();
                    swCheckApp2.Start();

                    Stopwatch swIsDiff = new Stopwatch();
                    Stopwatch swAppFill = new Stopwatch();

                    if (!noSync)
                    {
                        try
                        {
                            swIsDiff.Start();
                            bool isDiff = OutlookMgt.IsdifferentRecord(oAppointment, item);
                            swIsDiff.Stop();

                            if (isDiff)
                            {
                                if (log)
                                {
                                    serviceProxy.WriteInfo("++++++WebService: Update appointment from activity: " + item.ACTIVITYNUMBER);
                                }

                                if (!oAppointment.IsRecurring)
                                {
                                    swAppFill.Start();

                                    oAppointment.Subject = item.PURPOSE.ToString();
                                    oAppointment.Start = Convert.ToDateTime(item.STARTDATETIME.ToLocalTime());
                                    oAppointment.AllDayEvent = Convert.ToBoolean(item.ALLDAY);
                                    oAppointment.BillingInformation = item.BILLINGINFORMATION;
                                    oAppointment.Body = item.USERMEMO.ToString();
                                    oAppointment.Categories = item.OUTLOOKCATEGORIES;
                                    oAppointment.End = Convert.ToDateTime(item.ENDDATETIME.ToLocalTime());
                                    OutlookMgt.TASKPRIORITYConvertIntToOutlookOlImportance(item.TASKPRIORITY, ref oAppointment);
                                    oAppointment.Location = item.LOCATION;
                                    oAppointment.Mileage = item.MILEAGE;
                                    oAppointment.ReminderSet = Convert.ToBoolean(item.REMINDERACTIVE);
                                    oAppointment.ReminderMinutesBeforeStart = item.REMINDERMINUTES;
                                    oAppointment.Resources = item.OUTLOOKRESOURCES;
                                    oAppointment.ResponseRequested = Convert.ToBoolean(item.RESPONSEREQUESTED);

                                    OutlookMgt.SENSITIVITYIntToConvertOutlookOlSensitivity(item.SENSITIVITY, ref oAppointment);
                                    OutlookMgt.ACTIVITYTIMETYPEConvertIntToOutlookOlBusyStatus(item.ACTIVITYTIMETYPE, ref oAppointment);

                                    //B11796_OL2013CrmPluginScopeOfBusRels MGY 2015.11.13 Begin
                                    if (oAppointment.Body == null) oAppointment.Body = " ";
                                    oAppointment.Body = custDataInfo(oAppointment.Body, item);
                                    //B11796_OL2013CrmPluginScopeOfBusRels MGY 2015.11.13 End

                                    Outlook.UserProperty prop = oAppointment.UserProperties.Add("ActivityNum", Outlook.OlUserPropertyType.olText, false, Outlook.OlFormatText.olFormatTextText);
                                    prop.Value = item.ACTIVITYNUMBER;

                                    oAppointment.Save();
                                    swAppFill.Stop();
                                }

                                if (log)
                                {
                                    serviceProxy.WriteInfo("++++++WebService: Appointment saved.");
                                    serviceProxy.WriteInfo("++++++WebService: oAppointment Entry ID: " + oAppointment.EntryID);
                                    serviceProxy.WriteInfo("++++++WebService: CalFolder Store ID:" + calFolder.StoreID);
                                    printLogMsg("Activity saved in Outlook: " + item.ACTIVITYNUMBER);
                                }

                                if (isUpdateOutlookEntryID)
                                {
                                    serviceProxy.UpdateActivityOutlookEntryId(item.ACTIVITYNUMBER, company, oAppointment.EntryID);
                                }

                                if (insert)
                                {
                                    insertedOutlookItemCounter++;
                                }
                                else
                                {
                                    savedOutlookItemCounter++;
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            System.Windows.Forms.MessageBox.Show(e.Message);
                        }
                    }
                    swCheckApp2.Stop();
                        
                    if (!noAddList)
                    {
                        if (oAppointment.EntryID != null)
                        {
                            if (!smmActivityList.ContainsKey(oAppointment.EntryID))
                                smmActivityList.Add(oAppointment.EntryID, item.ACTIVITYNUMBER);

                        }
                    }

                    progressFromCrm.PerformStep();

                    swAct.Stop();
                    /*printLogMsg(String.Format("GetAppointment: {0, 4}; CheckApp: {1,4}; isDiff: {2,4}; AppFill: {3,4}; Activity process time: {4}",
                        swGetApp.ElapsedMilliseconds,
                        swCheckApp.ElapsedMilliseconds,
                        swIsDiff.ElapsedMilliseconds,
                        swAppFill.ElapsedMilliseconds,
                        swAct.ElapsedMilliseconds));*/

                } //end foreach
            }
            catch (InvalidOperationException ex)
            {
                printLogMsg(string.Format("Error transferring activities to Outlook! Error message: {0}", ex.Message));
            }
            catch (Exception ex)
            {
                printLogMsg(string.Format("Error transferring activities to Outlook! Error message: {0}", ex.Message));
            }

            String message = String.Empty;
            //if (updatedCRMItemCounter != 0)
            {
                message = message + "Number of activities updated in CRM: " + updatedCRMItemCounter + "\n";

            }
            //if (deletedCRMItemCounter != 0)
            {
                message = message + "Number of activities deleted from CRM: " + deletedCRMItemCounter + "\n";
            }

            //if (insertedOutlookItemCounter != 0)
            {
                message = message + "Number of activities transferred to Outlook: " + insertedOutlookItemCounter + "\n";

            }
            //if (savedOutlookItemCounter != 0)
            {
                message = message + "Number of activities updated in Outlook: " + savedOutlookItemCounter;
            }

            return message;
        }

        private String syncFromOutlookToCrm()
        {
            printLogMsg(string.Format("Synchronizing appointments from Outlook to CRM."));

            if (log)
                serviceProxy.WriteInfo("++++++WebService: syncFromCrmToOutlook");

            try
            {
                Stopwatch swGetFolder = new Stopwatch();
                swGetFolder.Start();
                Outlook.Folder calFolder = GetFolder(outlookCalendarID, outlookCalendarStoreID);
                swGetFolder.Stop();

                DateTime start = DateTime.Now.AddDays(-rangeDateBack);
                DateTime end = DateTime.Now.AddDays(rangeDateNext);

                Stopwatch swGetApps = new Stopwatch();
                swGetApps.Start();
                Outlook.Items rangeAppts = OutlookMgt.GetAppointmentsInRange(calFolder, start, end, this.lastSyncTime);
                swGetApps.Stop();

                Stopwatch swCntApps = new Stopwatch();
                swCntApps.Start();
                int apptCount = 0;

                //Ez a rész túl sokáig fut, holott csak számlálást végez. Megoldást kell keresni rá.
                if (rangeAppts != null)
                {
                    foreach (Outlook.AppointmentItem appt in rangeAppts)
                    {
                        ////B12773_CRMOutlookPluginIssues MGY 2016.07.04 Begin
                        //if (appt.LastModificationTime < this.lastSyncTime) continue;
                        ////B12773_CRMOutlookPluginIssues MGY 2016.07.04 End

                        if (appt.Sensitivity == Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate) continue;
                        //B19443_ChangesToPluginFunctionality MGY 2018.05.24 Begin
                        if (isPublicHoliday(appt)) continue;
                        //B19443_ChangesToPluginFunctionality MGY 2018.05.24 End

                        apptCount++;
                    }
                }
                //apptCount = 10;

                if (apptCount == 0)
                {
                    progressToCrm.Minimum = 0;
                    progressToCrm.Maximum = 1;
                    progressToCrm.Value = 0;
                    progressToCrm.Step = 1;
                    progressToCrm.PerformStep();
                }
                swCntApps.Stop();

                if (apptCount > 0)
                {
                    int taskProirity = 0;
                    int sensitivity = 0;
                    int activityTimeType = 0;
                    bool findItem = false;
                    string actnum;

                    progressToCrm.Minimum = 0;
                    progressToCrm.Maximum = apptCount;
                    progressToCrm.Value = 0;
                    progressToCrm.Step = 1;

                    foreach (Outlook.AppointmentItem appt in rangeAppts)
                    {
                        printLogMsg(string.Format("Checking Appointment: {0}.", appt.Subject));

                        //B12773_CRMOutlookPluginIssues MGY 2016.07.04 Begin
                        //Nem jó megoldás, mert amikor törlik az Activity párját AX-ból, akkor nem fog törlődni az appointment.
                        //if (appt.LastModificationTime < this.lastSyncTime) continue;
                        //B12773_CRMOutlookPluginIssues MGY 2016.07.04 End

                        //B08559_CRM_OutlookSyncBug MGY 2014.03.25 Begin
                        if (appt.Sensitivity == Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate) continue;
                        //B08559_CRM_OutlookSyncBug MGY 2014.03.25 End
                        //B19443_ChangesToPluginFunctionality MGY 2018.05.24 Begin
                        if (isPublicHoliday(appt)) continue;
                        //B19443_ChangesToPluginFunctionality MGY 2018.05.24 End

                        string partyID = string.Empty;
                        Outlook.UserProperties ups = appt.UserProperties;
                        Outlook.UserProperty prop = ups["PartyID"];
                        if (prop != null)
                        {
                            partyID = prop.Value;
                        }

                        findItem = false;
                        actnum = "";
                        var values = smmActivityList.Where(pair => pair.Key.Contains(appt.EntryID)).Select(pair => pair.Value);
                        foreach (var item in values)
                        {
                            findItem = true;
                        }

                        if (!findItem)
                        {
                            OutlookMgt.TASKPRIORITYConvertOutlookOlImportanceToInt(ref taskProirity, appt);
                            OutlookMgt.SENSITIVITYConvertOutlookOlSensitivityToInt(ref sensitivity, appt);
                            OutlookMgt.ACTIVITYTIMETYPEConvertOutlookOlBusyStatusToInt(ref activityTimeType, appt);

                            if (serviceProxy.IsDeletedActivity(appt.EntryID, company))
                            {
                                printLogMsg("Appointment removed as not found in CRM: " + appt.Subject);
                                appt.Delete();
                                deletedOutlookItemCounter++;
                                progressToCrm.PerformStep();
                                continue;
                            }

                            try
                            {
                                //Mi van akkor ha az activity korábban módosult mint az utolsó sync time, viszont az appointment utólag módosult?
                                //Ekkor duplán létrejön az Activity??
                                //Megoldás: Appointmentben letároljuk az activityNum-ot.
                                bool actUpdated = false;
                                Outlook.UserProperty uprops = appt.UserProperties.Find("ActivityNum");
                                if (uprops != null)
                                {
                                    actnum = Convert.ToString(uprops.Value);
                                    if (!String.IsNullOrEmpty(actnum))
                                    {
                                        serviceProxy.UpdateActivity(appt.Start.ToUniversalTime(),
                                                                appt.End.ToUniversalTime(),
                                                                empId,
                                                                appt.AllDayEvent,
                                                                appt.BillingInformation,
                                                                appt.Body,
                                                                appt.Subject,
                                                                appt.Categories,
                                                                taskProirity,
                                                                appt.Location,
                                                                appt.Mileage,
                                                                appt.ReminderSet,
                                                                appt.ReminderMinutesBeforeStart,
                                                                appt.Resources,
                                                                appt.ResponseRequested,
                                                                sensitivity,
                                                                activityTimeType,
                                                                networkAlias,
                                                                company,
                                                                appt.EntryID,
                                                                actnum);
                                        actUpdated = true;

                                    }
                                }

                                if (!actUpdated)
                                {
                                    actnum = serviceProxy.InsertActivity(appt.Start.ToUniversalTime(),
                                                                  appt.End.ToUniversalTime(),
                                                                  empId == null ? "" : empId,
                                                                  appt.AllDayEvent,
                                                                  appt.BillingInformation == null ? "" : appt.BillingInformation,
                                                                  appt.Body == null ? "" : appt.Body,
                                                                  appt.Subject == null ? "" : appt.Subject,
                                                                  appt.Categories == null ? "" : appt.Categories,
                                                                  taskProirity,
                                                                  appt.Location == null ? "" : appt.Location,
                                                                  appt.Mileage == null ? "" : appt.Mileage,
                                                                  appt.ReminderSet,
                                                                  appt.ReminderMinutesBeforeStart,
                                                                  appt.Resources == null ? "" : appt.Resources,
                                                                  appt.ResponseRequested,
                                                                  sensitivity,
                                                                  activityTimeType,
                                                                  networkAlias,
                                                                  company,
                                                                  appt.EntryID
                                                                  );

                                    Outlook.UserProperty propAppt = appt.UserProperties.Add("ActivityNum", Outlook.OlUserPropertyType.olText, false, Outlook.OlFormatText.olFormatTextText);
                                    propAppt.Value = actnum;

                                }

                                if (log)
                                {
                                    serviceProxy.WriteInfo("++++++WebService: " + "The " + actnum + " activity " + (actUpdated ? "updated!" : "inserted!"));
                                    printLogMsg(String.Format("Appointment '{0}' saved in CRM with identifier: {1}", appt.Subject, actnum));
                                }
                                savedCRMItemCounter++;
                            }
                            catch (Exception e)
                            {
                                System.Windows.Forms.MessageBox.Show("Could not insert activity: " + e.Message);
                            }
                        }
                        progressToCrm.PerformStep();
                    }
                }

                /*printLogMsg(String.Format("GetFolder: {0, 4}; GetApps: {1,4}; CountApps: {2,4}",
                    swGetFolder.ElapsedMilliseconds,
                    swGetApps.ElapsedMilliseconds,
                    swCntApps.ElapsedMilliseconds));*/

            }
            catch (Exception ex)
            {
                printLogMsg(string.Format("Error transferring appointments to CRM! Error message: {0}", ex.Message));
            }

            String message = String.Empty;
            //if (savedCRMItemCounter != 0)
            {
                message = message + "Number of activities transferred to CRM: " + savedCRMItemCounter + "\n";

            }

            //if (deletedOutlookItemCounter != 0)
            {
                message = message + "Number of activities deleted from Outlook: " + deletedOutlookItemCounter;
            }

            return message;
        }

        private void printLogMsg(String msg)
        {
            syncLogText.AppendText(msg);
            syncLogText.AppendText(System.Environment.NewLine);
            syncLogText.Refresh();
        }

        private void getNetworkCredentials()
        {
            WindowsIdentity wi = WindowsIdentity.GetCurrent();

            string domainUserName = wi.Name;

            networkDomain = domainUserName.Split('\\')[0] + ".Local";
            networkAlias = domainUserName.Split('\\')[1];
        }

        private bool chooseCalendar()
        {
            if (log)
                serviceProxy.WriteInfo("++++++WebService: string.IsNullOrEmpty(outlookCalendarID) || string.IsNullOrEmpty(outlookCalendarStoreID");

            Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");
            Outlook.MAPIFolder folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            List<DropDownItem> calendarItemList = new List<DropDownItem>();

            DropDownItem calendarItem = new DropDownItem(folder.Name, folder.EntryID, folder.StoreID);

            calendarItemList.Add(calendarItem);

            foreach (Outlook.MAPIFolder subFolder in folder.Folders)
            {
                DropDownItem calendarItemAdd = new DropDownItem(subFolder.Name, subFolder.EntryID, subFolder.StoreID);
                calendarItemList.Add(calendarItemAdd);
            }
            DropDownItem returnedItem = null;

            returnedItem = ShowDialog("Calendars:", "Please choose from calendars", calendarItemList);

            if (returnedItem == null)
            {
                System.Windows.Forms.MessageBox.Show("Plase choose calendar!");
                //throw new Exception("Please choose from calendars!");
                return false;
            }
            else
            {
                this.outlookCalendarID = returnedItem.OutlookCalendarID;
                this.outlookCalendarStoreID = returnedItem.OutlookCalendarStoreID;

                if (!string.IsNullOrEmpty(company)
                 && !string.IsNullOrEmpty(empId)
                 && !string.IsNullOrEmpty(this.outlookCalendarID)
                 && !string.IsNullOrEmpty(this.outlookCalendarStoreID))
                {
                    serviceProxy.UpdateEmplTable(empId, outlookCalendarID, outlookCalendarStoreID, company);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                    //throw new Exception("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                    return false;
                }
            }

            return true;
        }

        public DropDownItem ShowDialog(string text, string caption, List<DropDownItem> calendarList)
        {
            //
            Form prompt = new Form();
            prompt.Width = 250;
            prompt.Height = 150;
            prompt.Text = caption;
            Label textLabel = new Label() { Left = 20, Top = 10, Text = text };
            //TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Button confirmation = new Button() { Text = "Ok", Left = 20, Width = 50, Top = 80 };
            //Button cancel = new Button() { Text = "Cancel", Left = 75, Width = 50, Top = 80 };
            ComboBox comboBox = new ComboBox() { Left = 10, Width = 100, Top = 20 };
            comboBox.Location = new Point(50, 50);
            comboBox.Text = "";

            foreach (var item in calendarList)
            {
                comboBox.Items.Add(item.Name);
            }

            confirmation.Click += (sender, e) => { prompt.Close(); };
            //cancel.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(confirmation);
            //prompt.Controls.Add(cancel);
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(comboBox);

            prompt.ShowDialog();

            if (comboBox.SelectedItem == null)
            {
                return null;
            }
            foreach (var item in calendarList)
            {
                if (item.Name == comboBox.SelectedItem.ToString())
                {
                    return item;
                }
            }

            return null;
        }

        private Outlook.Folder GetFolder(string _outlookCalendarID, string _outlookCalendarStoreID)
        {
            try
            {
                //Outlook.Application outlookApp = new Outlook.Application();
                return (Outlook.Folder)outlookApp.Session.GetFolderFromID(_outlookCalendarID, _outlookCalendarStoreID);

            }
            catch (Exception x)
            {

                return null;
            }
        }

        private Outlook.AppointmentItem GetAppointment(string _appointmentEntryID, string _calendarStoreID)
        {
            try
            {
                Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI"); // <-- Made it global, 2016.06.29 MGY
                return (Outlook.AppointmentItem)ns.GetItemFromID(_appointmentEntryID, _calendarStoreID);
            }
            catch (Exception x)
            {
                return null;
            }
        }

        private String custDataInfo(String bodyText, Model.ActivityData item)
        {
            if (bodyText == null) bodyText = String.Empty;
            String custDataBegin = "## CRM Data ==> ##";
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
            custDataInfo += String.Format("Customer name: {0}", item.CUSTNAME) + Environment.NewLine;
            custDataInfo += String.Format("Customer ABC: {0}", item.CUSTABC) + Environment.NewLine;
            custDataInfo += String.Format("Sales Rep: {0}", item.CUSTSALESREP) + Environment.NewLine;
            custDataInfo += String.Format("Mode of delivery: {0}", item.DLVMODE) + Environment.NewLine;
            custDataInfo += String.Format("Search name: {0}", item.NAMEALIAS) + Environment.NewLine;
            custDataInfo += String.Format("Contact ID: {0}", item.CONTACTPERSONID) + Environment.NewLine;
            custDataInfo += String.Format("Contact name: {0}", item.CONTACTNAME) + Environment.NewLine;
            custDataInfo += String.Format("Contact phone : {0}", item.CONTACTPHONE) + Environment.NewLine;
            custDataInfo += String.Format("Contact email : {0}", item.CONTACTEMAIL) + Environment.NewLine;
            custDataInfo += String.Format("Activity number: {0}", item.ACTIVITYNUMBER);

            custDataTxt = Environment.NewLine;
            custDataTxt += custDataBegin + Environment.NewLine;
            custDataTxt += custDataInfo + Environment.NewLine;
            custDataTxt += custDataEnd + Environment.NewLine;

            bodyText = bodyText.Insert(cdStartPos, custDataTxt);

            return bodyText;
        }

        //B19443_ChangesToPluginFunctionality MGY 2018.05.24
        private bool isPublicHoliday(Outlook.AppointmentItem appt)
        {

            CountryCode countryCode = CountryCode.NL;
            switch (this.CompanyCountry)
            {
                case "NL":
                    countryCode = CountryCode.NL;
                    break;
                case "BE":
                    countryCode = CountryCode.BE;
                    break;
                case "DK":
                    countryCode = CountryCode.DK;
                    break;
                case "ES":
                    countryCode = CountryCode.ES;
                    break;
                case "DE":
                    countryCode = CountryCode.DE;
                    break;
                default:
                    countryCode = CountryCode.NL;
                    break;
            }

            return (DateSystem.IsPublicHoliday(appt.Start, countryCode) ||
                (appt.Start.DayOfWeek == DayOfWeek.Saturday) ||
                (appt.Start.DayOfWeek == DayOfWeek.Sunday));
        }

    }

}
