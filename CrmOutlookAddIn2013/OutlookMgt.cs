using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Drawing;
using System.Net;
using System.Security.Principal;

namespace CrmOutlookAddIn2013
{
    public class OutlookMgt
    {
        WebOutlookCrm.OutlookService service = new WebOutlookCrm.OutlookService();


        Outlook.Application outlookApp = new Outlook.Application();        

        public string outlookCalendarID { get; set; }
        public string outlookCalendarStoreID { get; set; }
        public string empId { get; private set; }
        public string comapny { get; set; }
        public int rangeDateBack { get; set; }
        public int rangeDateNext { get; set; }
        public bool log { get; set; }

        public int deletedCRMItemCounter        = 0;
        public int updatedCRMItemCounter        = 0;
        public int savedCRMItemCounter          = 0;
        public int insertedOutlookItemCounter   = 0;
        public int deletedOutlookItemCounter    = 0;
        public int savedOutlookItemCounter      = 0;

        Dictionary<string, string> smmActivitiList = new Dictionary<string, string>();

        string networkDomain;
        string networkAlias;

        public OutlookMgt() //Konstruktor
        {
            
        }
        
        public bool Sync()
        {           
            service.UseDefaultCredentials = true;
            this.log = service.IsLog();                        
                                    
            WindowsIdentity wi = WindowsIdentity.GetCurrent();

            string domainUserName = wi.Name;

            networkDomain = domainUserName.Split('\\')[0] + ".Local";
            networkAlias = domainUserName.Split('\\')[1];

            if (service.IsTestMode())
            {
                networkAlias = service.GetNetworkAlias();
            }

            WebOutlookCrm.smmActivities.USERDATADataTable userDataTable = service.GetUserData(networkDomain, networkAlias);

            
            
            //Csak egy sor lesz
            if (userDataTable.Count != 0)
            {
                //foreach (WebOutlookCrm.smmActivities.USERDATADataTable item in userDataTable)
               // {
                if (log)
                    service.WriteInfo("++++++WebService: userDataTable.Count != 0");
                this.outlookCalendarID = userDataTable[0].OUTLOOKCALENDAROUTLOOKENTRYID;

                this.outlookCalendarStoreID = (userDataTable[0].OUTLOOKCALENDAROUTLOOKSTOREID == null) ? string.Empty : userDataTable[0].OUTLOOKCALENDAROUTLOOKSTOREID.ToString();
                this.comapny = userDataTable[0].COMPANY;
                this.empId = userDataTable[0].EMPLID;                    
                //}

                this.rangeDateBack = service.FilterDateBack();
                this.rangeDateNext = service.FilterDateNext();

                if (string.IsNullOrEmpty(outlookCalendarID) || string.IsNullOrEmpty(outlookCalendarStoreID))
                {
                    if (log)
                        service.WriteInfo("++++++WebService: string.IsNullOrEmpty(outlookCalendarID) || string.IsNullOrEmpty(outlookCalendarStoreID");
                    Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");
                    Outlook.MAPIFolder folder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

                    List<DropDownItem> calendarItemList = new List<DropDownItem>();

                    DropDownItem calendarItem = new DropDownItem(folder.Name,folder.EntryID,folder.StoreID);

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
                        ///System.Windows.Forms.MessageBox.Show("Plase chose calendar!");
                        throw new Exception("Please choose from calendars!");
                        //return false;
                    }
                    else
                    {
                        this.outlookCalendarID = returnedItem.OutlookCalendarID;
                        this.outlookCalendarStoreID = returnedItem.OutlookCalendarStoreID;

                        if (!string.IsNullOrEmpty(comapny) && !string.IsNullOrEmpty(empId) && !string.IsNullOrEmpty(this.outlookCalendarID) && !string.IsNullOrEmpty(this.outlookCalendarStoreID))
                        {                           
                            service.UpdateEmplTable(empId, outlookCalendarID, outlookCalendarStoreID, comapny);
                        }
                        else
                        {
                            //System.Windows.Forms.MessageBox.Show("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                            throw new Exception("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                            //return false;
                        }
                    }
                    
                }

                if (!string.IsNullOrEmpty(comapny) && !string.IsNullOrEmpty(empId) && !string.IsNullOrEmpty(this.outlookCalendarID) && !string.IsNullOrEmpty(this.outlookCalendarStoreID))
                {
                    if (log)
                        service.WriteInfo("++++++WebService: SetAppointment");
                    SetAppointment(comapny, empId);
                    smmActivitiList = null;
                }
                else
                {
                    //System.Windows.Forms.MessageBox.Show("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID)!");
                    //return false;
                    throw new Exception("The user data table is not set (empty empId or company or OutlookCalendarStoreID or  OutlookCalendarID, DoaminUser: " + networkAlias + "Domain: " + networkDomain + ")!");
                }
                
            }
            else
            {
                //System.Windows.Forms.MessageBox.Show("The user data table is not set!");
                //return false;
                throw new Exception("The user data table is not set!");
            }
            
           
            return true; ;
            //SetAppointment("BDT", "LWIEGAARD");
           
        }

        

        public void SetAppointment(string _company, string _responsibleEployee)
        {
            bool isUpdateOutlookEntryID = false;
            int taskProirity = 0;
            int sensitivity = 0;
            int activityTimeType = 0;
            bool noSync = false;
            bool noAddList = false;

            deletedCRMItemCounter = 0;
            updatedCRMItemCounter = 0;
            savedCRMItemCounter = 0;

            insertedOutlookItemCounter = 0;
            deletedOutlookItemCounter = 0;

            WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable smmAct = service.GetsmmActivities(_company, _responsibleEployee);
           
            Outlook.AppointmentItem oAppointment = null;
            Outlook.Folder calFolder = null;

            //Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");
            //Outlook.MAPIFolder calFolder = ns.GetFolderFromID(outlookCalendarID, outlookCalendarStoreID);
                        

            calFolder = GetFolder(outlookCalendarID, outlookCalendarStoreID);

            //Outlook.MAPIFolder folder = calFolder;                      

            foreach (WebOutlookCrm.smmActivities.SMMACTIVITIESRow item in smmAct)
            {
                if (log)
                {
                    service.WriteInfo("++++++WebService: foreach SetAppointment");
                    service.WriteInfo("++++++WebService: ActNumber: "+item.ACTIVITYNUMBER);
                }
                noSync = false;
                noAddList = false;                
                oAppointment = null;
                bool insert = false;
                oAppointment = GetAppointment(item.OUTLOOKENTRYID, outlookCalendarStoreID); 
                              
                if (oAppointment == null)
                {
                    
                    if (log)
                    {
                        service.WriteInfo("++++++WebService: oAppointment == null");
                        service.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                    }

                    if (string.IsNullOrEmpty(item.OUTLOOKENTRYID))
                    {
                        
                        isUpdateOutlookEntryID = true;
                        insert = true;
                        oAppointment = (Outlook.AppointmentItem)calFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);

                        if (log)
                        {
                            service.WriteInfo("++++++WebService: string.IsNullOrEmpty(item.OUTLOOKENTRYID)");
                            service.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                            service.WriteInfo("++++++WebService: noSync: " + noSync);
                        }

                        //oAppointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);                        
                    }
                    else
                    {
                        // ez az ág mikor nem találja az outlookban viszont van OutlikEntryID-ja tehát az outlookból lett törölve
                        if (log)
                        {
                            service.WriteInfo("++++++WebService: string.IsNullOrEmpty(item.OUTLOOKENTRYID) != ");
                            service.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                        }

                        noSync = true;
                        insert = false;
                        noAddList = service.DeleteActivity(item.OUTLOOKENTRYID, comapny);
                        //System.Windows.Forms.MessageBox.Show("The " + item.ACTIVITYNUMBER + " activities deleted!");                        
                        if (log)
                        {
                            service.WriteInfo("++++++WebService: "+"The " + item.ACTIVITYNUMBER + " activities deleted!");                            
                        }
                        deletedCRMItemCounter++;
                    }
                }
                else
                {
                    insert = false;

                    if (log)
                    {
                        service.WriteInfo("++++++WebService: oAppointment != null");
                        service.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                    }

                    isUpdateOutlookEntryID = false;

                    if (oAppointment.LastModificationTime > item.MODIFIEDDATETIME.ToLocalTime())
                    {
                        if (IsdifferentRecord(oAppointment, item))
                        {
                            if (log)
                            {
                                service.WriteInfo("++++++WebService: oAppointment.LastModificationTime > item.MODIFIEDDATETIME.ToLocalTime()");
                                service.WriteInfo("++++++WebService: " + oAppointment.LastModificationTime + " " + item.MODIFIEDDATETIME.ToLocalTime());
                                service.WriteInfo("++++++WebService: ActNumber: " + item.ACTIVITYNUMBER);
                            }

                            taskProirity = 0;
                            sensitivity = 0;
                            activityTimeType = 0;

                            OutlookMgt.TASKPRIORITYConvertOutlookOlImportanceToInt(ref taskProirity, oAppointment);
                            OutlookMgt.SENSITIVITYConvertOutlookOlSensitivityToInt(ref sensitivity, oAppointment);
                            OutlookMgt.ACTIVITYTIMETYPEConvertOutlookOlBusyStatusToInt(ref activityTimeType, oAppointment);

                            noSync = service.UpdateActivity(oAppointment.Start.ToUniversalTime(),
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
                                                            comapny,
                                                            oAppointment.EntryID,
                                                            item.ACTIVITYNUMBER);                            
                            if (log)
                            {
                                service.WriteInfo("++++++WebService: " + "The " + item.ACTIVITYNUMBER + " activities updated!");
                            }
                            updatedCRMItemCounter++;
                        }
                      
                    }                    
                }                                
            
                if (!noSync)
                {
                    try
                    {
                        if (IsdifferentRecord(oAppointment, item))
                        {                                                             
                            if (log)
                            {
                                service.WriteInfo("++++++WebService: Create new appointment from activity: " + item.ACTIVITYNUMBER);
                            }

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

                            oAppointment.Save();

                            if (log)
                            {
                                service.WriteInfo("++++++WebService: Appointment saved.");
                                service.WriteInfo("++++++WebService: oAppointment Entry ID: " + oAppointment.EntryID);
                                service.WriteInfo("++++++WebService: CalFolder Store ID:" + calFolder.StoreID);
                            }
                            if (isUpdateOutlookEntryID)
                            {
                                service.UpdateActivityOutlookEntryId(item.ACTIVITYNUMBER, comapny, oAppointment.EntryID);                            
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

                if (!noAddList)
                {
                    if (oAppointment.EntryID != null)
                    {
                        smmActivitiList.Add(oAppointment.EntryID, item.ACTIVITYNUMBER);
                    }
                }
            }

            this.SearchAppointments(calFolder);
        }

        public static Boolean IsdifferentRecord(Outlook.AppointmentItem oAppointment, WebOutlookCrm.smmActivities.SMMACTIVITIESRow item)
        {
            Boolean differenetRec = false;

            if (oAppointment.Start != item.STARTDATETIME.ToLocalTime())
            {
                differenetRec = true;
            }
            if (oAppointment.End != item.ENDDATETIME.ToLocalTime())
            {
                differenetRec = true;
            }


            if (string.IsNullOrEmpty(oAppointment.BillingInformation))
            {

                if (string.IsNullOrEmpty(item.BILLINGINFORMATION))
                {

                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.BillingInformation != item.BILLINGINFORMATION)
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Body))
            {
                if (string.IsNullOrEmpty(item.USERMEMO))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Body.ToString() != item.USERMEMO)
                {
                    differenetRec = true;
                }
            }

            if (string.IsNullOrEmpty(oAppointment.Subject))
            {
                if (string.IsNullOrEmpty(item.PURPOSE))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Subject.ToString() != item.PURPOSE.ToString())
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Categories))
            {
                if (string.IsNullOrEmpty(item.OUTLOOKCATEGORIES))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Categories.ToString() != item.OUTLOOKCATEGORIES)
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Location))
            {
                if (string.IsNullOrEmpty(item.LOCATION))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Location.ToString() != item.LOCATION.ToString())
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Mileage))
            {
                if (string.IsNullOrEmpty(item.MILEAGE))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Mileage.ToString() != item.MILEAGE.ToString())
                {
                    differenetRec = true;
                }
            }

            if (oAppointment.ReminderMinutesBeforeStart != item.REMINDERMINUTES)
            {
                differenetRec = true;
            }
            if (string.IsNullOrEmpty(oAppointment.Resources))
            {
                if (string.IsNullOrEmpty(item.OUTLOOKRESOURCES))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Resources.ToString() != item.OUTLOOKRESOURCES)
                {
                    differenetRec = true;
                }
            }
            if (oAppointment.AllDayEvent != Convert.ToBoolean(item.ALLDAY))
            {
                differenetRec = true;
            }

            if (oAppointment.ReminderSet != Convert.ToBoolean(item.REMINDERACTIVE))
            {
                differenetRec = true;
            }

            if (oAppointment.ResponseRequested != Convert.ToBoolean(item.RESPONSEREQUESTED))
            {
                differenetRec = true;
            }
            if (oAppointment.Importance != OutlookMgt.TASKPRIORITYConvertIntToOutlookOlImportance(item.TASKPRIORITY))
            {
                differenetRec = true;
            }
            if (oAppointment.Sensitivity != OutlookMgt.SENSITIVITYIntToConvertOutlookOlSensitivity(item.SENSITIVITY))
            {
                differenetRec = true;
            }
            if (oAppointment.BusyStatus != OutlookMgt.ACTIVITYTIMETYPEConvertIntToOutlookOlBusyStatus(item.ACTIVITYTIMETYPE))
            {
                differenetRec = true;
            }
            return differenetRec;
        }

        public static Boolean IsdifferentRecord(Outlook.AppointmentItem oAppointment, Model.ActivityData item)
        {
            Boolean differenetRec = false;

            if (oAppointment.Start != item.STARTDATETIME.ToLocalTime())
            {
                differenetRec = true;
            }
            if (oAppointment.End != item.ENDDATETIME.ToLocalTime())
            {
                differenetRec = true;
            }


            if (string.IsNullOrEmpty(oAppointment.BillingInformation))
            {

                if (string.IsNullOrEmpty(item.BILLINGINFORMATION))
                {

                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.BillingInformation != item.BILLINGINFORMATION)
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Body))
            {
                if (string.IsNullOrEmpty(item.USERMEMO))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Body.ToString() != item.USERMEMO)
                {
                    differenetRec = true;
                }
            }

            if (string.IsNullOrEmpty(oAppointment.Subject))
            {
                if (string.IsNullOrEmpty(item.PURPOSE))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Subject.ToString() != item.PURPOSE.ToString())
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Categories))
            {
                if (string.IsNullOrEmpty(item.OUTLOOKCATEGORIES))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Categories.ToString() != item.OUTLOOKCATEGORIES)
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Location))
            {
                if (string.IsNullOrEmpty(item.LOCATION))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Location.ToString() != item.LOCATION.ToString())
                {
                    differenetRec = true;
                }
            }
            if (string.IsNullOrEmpty(oAppointment.Mileage))
            {
                if (string.IsNullOrEmpty(item.MILEAGE))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Mileage.ToString() != item.MILEAGE.ToString())
                {
                    differenetRec = true;
                }
            }

            if (oAppointment.ReminderMinutesBeforeStart != item.REMINDERMINUTES)
            {
                differenetRec = true;
            }
            if (string.IsNullOrEmpty(oAppointment.Resources))
            {
                if (string.IsNullOrEmpty(item.OUTLOOKRESOURCES))
                {
                }
                else
                {
                    differenetRec = true;
                }
            }
            else
            {
                if (oAppointment.Resources.ToString() != item.OUTLOOKRESOURCES)
                {
                    differenetRec = true;
                }
            }
            if (oAppointment.AllDayEvent != Convert.ToBoolean(item.ALLDAY))
            {
                differenetRec = true;
            }

            if (oAppointment.ReminderSet != Convert.ToBoolean(item.REMINDERACTIVE))
            {
                differenetRec = true;
            }

            if (oAppointment.ResponseRequested != Convert.ToBoolean(item.RESPONSEREQUESTED))
            {
                differenetRec = true;
            }
            if (oAppointment.Importance != OutlookMgt.TASKPRIORITYConvertIntToOutlookOlImportance(item.TASKPRIORITY))
            {
                differenetRec = true;
            }
            if (oAppointment.Sensitivity != OutlookMgt.SENSITIVITYIntToConvertOutlookOlSensitivity(item.SENSITIVITY))
            {
                differenetRec = true;
            }
            if (oAppointment.BusyStatus != OutlookMgt.ACTIVITYTIMETYPEConvertIntToOutlookOlBusyStatus(item.ACTIVITYTIMETYPE))
            {
                differenetRec = true;
            }
            return differenetRec;
        }

        void _CalendarItems_ItemAdd(object Item)
        {
            throw new NotImplementedException();                        
        }

        private Outlook.AppointmentItem GetAppointment(string _appointmentEntryID, string _calendarStoreID)
        {
            try
            {
                //Outlook.Application outlookApp = new Outlook.Application();
                Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");
                return (Outlook.AppointmentItem)ns.GetItemFromID(_appointmentEntryID, _calendarStoreID);
            }
            catch (Exception x)
            {
                return null;
            }
        }
        private Outlook.Folder GetFolder(string _outlookCalendarID,string  _outlookCalendarStoreID)
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

        //Ez a metodus keresi a bejegyzeseket Outlookban.
        private void SearchAppointments(Outlook.Folder _calFolder)
        {
            
            DateTime start = DateTime.Now.AddDays(-rangeDateBack);
            DateTime end = DateTime.Now.AddDays(rangeDateNext);
            Outlook.Items rangeAppts = GetAppointmentsInRange(_calFolder, start, end, DateTime.MinValue);
            

            if (rangeAppts != null)
            {
                int taskProirity = 0;
                int sensitivity = 0;
                int activityTimeType = 0;
                bool findItem = false;
                string actnum;                

                foreach (Outlook.AppointmentItem appt in rangeAppts)
                {
                    //B08559_CRM_OutlookSyncBug MGY 2014.03.25 Begin
                    if (appt.Sensitivity == Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate) continue;
                    //B08559_CRM_OutlookSyncBug MGY 2014.03.25 End

                    findItem = false;
                    actnum = "";
                    var values = smmActivitiList.Where(pair => pair.Key.Contains(appt.EntryID)).Select(pair => pair.Value);
                    foreach (var item in values)
	                {
                        findItem = true;                        		 
                    }

                    if (!findItem)
                    {
                        OutlookMgt.TASKPRIORITYConvertOutlookOlImportanceToInt(ref taskProirity, appt);
                        OutlookMgt.SENSITIVITYConvertOutlookOlSensitivityToInt(ref sensitivity, appt);
                        OutlookMgt.ACTIVITYTIMETYPEConvertOutlookOlBusyStatusToInt(ref activityTimeType, appt);

                        if (service.IsDeletedActivity(appt.EntryID, comapny))
                        {
                            appt.Delete();
                            deletedOutlookItemCounter++;
                            continue;
                        }
                                          
                          actnum =   service.InsertsmmActivities(appt.Start.ToUniversalTime(),
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
                                                        comapny,
                                                        appt.EntryID
                                                        );

                          //System.Windows.Forms.MessageBox.Show("The " + actnum + " activities inserted!");
                          if (log)
                          {
                              service.WriteInfo("++++++WebService: " + "The " + actnum + " activities inserted!");                             
                          }
                          savedCRMItemCounter++;
                    }                 
                }

                //System.Windows.Forms.MessageBox.Show("The " + insertedItemCounter + " activities inserted in Ax!");
            }
             
        }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns>Outlook.Items</returns>
        public static Outlook.Items GetAppointmentsInRange(Outlook.Folder folder, DateTime startTime, DateTime endTime, DateTime lastSync)
        {

            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";
                //B08559_CRM_Outlook_sync_bug MGY 2014.03.14 Begin
                //+ " AND [Sensitivity] <> " + Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate;
                //B08559_CRM_Outlook_sync_bug MGY 2014.03.14 End
                //B12773_CRMOutlookPluginIssues MGY 2016.07.04 Begin
                //Ez nem működik IncludeRecurrences = true esetén:
                //+ " AND [LastModificationTime] > '" + lastSync.ToString("g") + "'";
                //B12773_CRMOutlookPluginIssues MGY 2016.07.04 End

            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

        public static void TASKPRIORITYConvertIntToOutlookOlImportance(int _tASKPRIORITY,ref Outlook.AppointmentItem _oAppointment)
        {
            
            switch (_tASKPRIORITY)
                {
                    case 0:
                        _oAppointment.Importance = Outlook.OlImportance.olImportanceLow;
                        break;
                    case 1:
                        _oAppointment.Importance = Outlook.OlImportance.olImportanceNormal;
                        break;
                    case 2:
                        _oAppointment.Importance = Outlook.OlImportance.olImportanceHigh;
                        break;                    
                }
        }

        public static Outlook.OlImportance TASKPRIORITYConvertIntToOutlookOlImportance(int _tASKPRIORITY)
        {
            Outlook.OlImportance importance = Outlook.OlImportance.olImportanceHigh;

            switch (_tASKPRIORITY)
            {
                case 0:
                    importance = Outlook.OlImportance.olImportanceLow;
                    break;
                case 1:
                    importance = Outlook.OlImportance.olImportanceNormal;
                    break;
                case 2:
                    importance = Outlook.OlImportance.olImportanceHigh;
                    break;
            }

            return importance;
        }

        public static void TASKPRIORITYConvertOutlookOlImportanceToInt(ref int _tASKPRIORITY, Outlook.AppointmentItem _oAppointment)
        {
            switch (_oAppointment.Importance)
            {
                case Outlook.OlImportance.olImportanceLow:
                    _tASKPRIORITY = 0;
                    break;
                case Outlook.OlImportance.olImportanceNormal:
                    _tASKPRIORITY = 1;
                    break;
                case Outlook.OlImportance.olImportanceHigh:
                    _tASKPRIORITY = 2;
                    break;
            }
        }

        public static void SENSITIVITYIntToConvertOutlookOlSensitivity(int _sENSITIVITY, ref Outlook.AppointmentItem _oAppointment)
        {
            switch (_sENSITIVITY)
                {
                    case 0:
                        _oAppointment.Sensitivity = Outlook.OlSensitivity.olNormal;
                        break;
                    case 1:
                        _oAppointment.Sensitivity = Outlook.OlSensitivity.olPersonal;
                        break;
                    case 2:
                        _oAppointment.Sensitivity = Outlook.OlSensitivity.olPrivate;
                        break;
                    case 3:
                        _oAppointment.Sensitivity = Outlook.OlSensitivity.olConfidential;
                        break;                 
                }
        }

        public static Outlook.OlSensitivity SENSITIVITYIntToConvertOutlookOlSensitivity(int _sENSITIVITY)
        {
            Outlook.OlSensitivity sensitivity = Outlook.OlSensitivity.olNormal;

            switch (_sENSITIVITY)
            {
                case 0:
                    sensitivity = Outlook.OlSensitivity.olNormal;
                    break;
                case 1:
                    sensitivity = Outlook.OlSensitivity.olPersonal;
                    break;
                case 2:
                    sensitivity = Outlook.OlSensitivity.olPrivate;
                    break;
                case 3:
                    sensitivity = Outlook.OlSensitivity.olConfidential;
                    break;
            }

            return sensitivity;
        }


        public static void SENSITIVITYConvertOutlookOlSensitivityToInt(ref int _sENSITIVITY,Outlook.AppointmentItem _oAppointment)
        {
            switch (_oAppointment.Sensitivity)
            {
                case Outlook.OlSensitivity.olNormal:                    
                    _sENSITIVITY = 0;
                    break;
                case Outlook.OlSensitivity.olPersonal:                    
                    _sENSITIVITY = 1;
                    break;
                case Outlook.OlSensitivity.olPrivate:                    
                    _sENSITIVITY = 2;
                    break;
                case Outlook.OlSensitivity.olConfidential:
                    _sENSITIVITY = 3;                    
                    break;
            }
        }
        public static void ACTIVITYTIMETYPEConvertIntToOutlookOlBusyStatus(int _aCTIVITYTIMETYPE,ref Outlook.AppointmentItem _oAppointment)
        {
            switch (_aCTIVITYTIMETYPE)
            {
                case 0:
                    _oAppointment.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree;
                    break;
                case 1:
                    _oAppointment.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olTentative;
                    break;
                case 2:
                    _oAppointment.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy;
                    break;
                case 3:
                    _oAppointment.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olOutOfOffice;
                    break;               
            }
        }

        public static Microsoft.Office.Interop.Outlook.OlBusyStatus ACTIVITYTIMETYPEConvertIntToOutlookOlBusyStatus(int _aCTIVITYTIMETYPE)
        {
            Microsoft.Office.Interop.Outlook.OlBusyStatus busyStatus;
            switch (_aCTIVITYTIMETYPE)
            {
                case 0:
                    busyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree;
                    break;
                case 1:
                    busyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olTentative;
                    break;
                case 2:
                    busyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy;
                    break;
                case 3:
                    busyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olOutOfOffice;
                    break;
                default:
                    busyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy;
                    break;               
            }

            return busyStatus;

        }

        public static void ACTIVITYTIMETYPEConvertOutlookOlBusyStatusToInt(ref int _aCTIVITYTIMETYPE, Outlook.AppointmentItem _oAppointment)
        {
            switch (_oAppointment.BusyStatus)
            {
                    
                case Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree:
                    
                    _aCTIVITYTIMETYPE = 0;
                    break;
                case Microsoft.Office.Interop.Outlook.OlBusyStatus.olTentative:                    
                    _aCTIVITYTIMETYPE = 1;
                    break;
                case Microsoft.Office.Interop.Outlook.OlBusyStatus.olBusy:                    
                    _aCTIVITYTIMETYPE = 2;
                    break;
                case Microsoft.Office.Interop.Outlook.OlBusyStatus.olOutOfOffice:                   
                    _aCTIVITYTIMETYPE = 3;
                    break;
            }
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
            ComboBox comboBox = new ComboBox() {Left = 10, Width = 100, Top = 20 }; 
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
     
    }


    public class DropDownItem
    {
        public string Name = "";
        public string OutlookCalendarID = "";
        public string OutlookCalendarStoreID = "";

        public DropDownItem(string name, string outlookCalendarID, string outlookCalendarStoreID)
        {
            this.Name = name;
            this.OutlookCalendarID = outlookCalendarID;
            this.OutlookCalendarStoreID = outlookCalendarStoreID; 
        }       
    }
}
