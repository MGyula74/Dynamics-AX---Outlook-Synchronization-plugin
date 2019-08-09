using CrmOutlookAddIn2013.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CrmOutlookAddIn2013
{
    class CrmWebServiceProxy
    {
        public bool isTest { get; set; }
        UserData userData = null;
        WebOutlookCrm.OutlookService serviceLIV = new WebOutlookCrm.OutlookService();
        WebOutlookCrmST.OutlookService serviceSTG = new WebOutlookCrmST.OutlookService();

        public CrmWebServiceProxy(string networkDomain, string networkAlias)
        {
            this.GetUserData(networkDomain, networkAlias);
        }

        public bool UseDefaultCredentials
        {
            set
            {
                if (isTest)
                    serviceSTG.UseDefaultCredentials = value;
                else
                    serviceLIV.UseDefaultCredentials = value;
            }
        }

        public string Url
        {
            get
            {
                return isTest ? serviceSTG.Url : serviceLIV.Url;
            }
        }

        public bool IsLog()
        {
            return isTest ? serviceSTG.IsLog() : serviceLIV.IsLog();
        }

        public bool IsTestMode()
        {
            return isTest ? serviceSTG.IsTestMode() : serviceLIV.IsTestMode();
        }

        public string GetTestNetworkAlias()
        {
            return isTest ? serviceSTG.GetNetworkAlias() : serviceLIV.GetNetworkAlias();
        }

        public UserData GetUserData(string networkDomain, string networkAlias)
        {
            if (userData != null && userData.networkDomain == networkDomain && userData.networkAlias == networkAlias)
                return userData;
            
            serviceLIV.UseDefaultCredentials = true;
            WebOutlookCrm.smmActivities.USERDATADataTable userDataTable = serviceLIV.GetUserData(networkDomain, networkAlias);
            if (userDataTable.Count != 0)
            {
                userData = new UserData();
                userData.OUTLOOKCALENDAROUTLOOKENTRYID = userDataTable[0].OUTLOOKCALENDAROUTLOOKENTRYID;
                userData.OUTLOOKCALENDAROUTLOOKSTOREID = userDataTable[0].OUTLOOKCALENDAROUTLOOKSTOREID;
                userData.COMPANY = userDataTable[0].COMPANY;
                userData.EMPLID = userDataTable[0].EMPLID;
                userData.smmSynchronizeDaysBack = userDataTable[0].smmSynchronizeDaysBack;
                userData.smmSynchronizeDaysForward = userDataTable[0].smmSynchronizeDaysForward;
                userData.CompanyCountryId = userDataTable[0].LANGUAGEID;
                isTest = false;
            }
            else
            {
                serviceSTG.UseDefaultCredentials = true;
                WebOutlookCrmST.smmActivities.USERDATADataTable userDataTableST = serviceSTG.GetUserData(networkDomain, networkAlias);
                if (userDataTableST.Count == 0)
                    return null;

                userData = new UserData();
                userData.OUTLOOKCALENDAROUTLOOKENTRYID = userDataTableST[0].OUTLOOKCALENDAROUTLOOKENTRYID;
                userData.OUTLOOKCALENDAROUTLOOKSTOREID = userDataTableST[0].OUTLOOKCALENDAROUTLOOKSTOREID;
                userData.COMPANY = userDataTableST[0].COMPANY;
                userData.EMPLID = userDataTableST[0].EMPLID;
                userData.smmSynchronizeDaysBack = userDataTableST[0].smmSynchronizeDaysBack;
                userData.smmSynchronizeDaysForward = userDataTableST[0].smmSynchronizeDaysForward;
                userData.CompanyCountryId = userDataTableST[0].LANGUAGEID;
                isTest = true;
            }

            userData.networkDomain = networkDomain;
            userData.networkAlias = networkAlias;

            return userData;
        }

        public bool UpdateEmplTable(string emplId, string OUTLOOKENTRYID, string OUTLOOKCALENDARSTOREID, string company)
        {
            return isTest ? serviceSTG.UpdateEmplTable(emplId, OUTLOOKENTRYID, OUTLOOKCALENDARSTOREID, company) :
                serviceLIV.UpdateEmplTable(emplId, OUTLOOKENTRYID, OUTLOOKCALENDARSTOREID, company);
        }

        //public USERDATADataTable GetUserData(string networkDomain, string networkAlias)
        //{
        //    if (isTest)
        //    {
        //        WebOutlookCrmST.smmActivities.USERDATADataTable userDataTable = serviceSTG.GetUserData(networkDomain, networkAlias);
        //        foreach (WebOutlookCrmST.smmActivities.USERDATARow row in userDataTable)
        //        {
        //            USERDATADataRow rowOut = new USERDATADataRow();
        //            rowOut.OUTLOOKCALENDAROUTLOOKENTRYID = row.OUTLOOKCALENDAROUTLOOKENTRYID;
        //        }
        //    }
        //}
        public void WriteInfo(string msg)
        {
            if (isTest) serviceSTG.WriteInfo(msg); else serviceLIV.WriteInfo(msg);
        }

        public int FilterDateBack()
        {
            return isTest ? serviceSTG.FilterDateBack() : serviceLIV.FilterDateBack();
        }

        public int FilterDateNext()
        {
            return isTest ? serviceSTG.FilterDateNext() : serviceLIV.FilterDateNext();
        }

        public DateTime getLastSynchronizationTime(string company, string emplId)
        {
            return isTest ? serviceSTG.getLastSynchronizationTime(company, emplId) : serviceLIV.getLastSynchronizationTime(company, emplId);
        }

        public void synchronizationFinished(string company, string empId, DateTime clientTime)
        {
            if (isTest)
                serviceSTG.synchronizationFinished(company, empId, clientTime);
            else
                serviceLIV.synchronizationFinished(company, empId, clientTime);
        }

        public List<ActivityData> GetsmmActivities(string company, string responsible)
        {
            List<ActivityData> actList = new List<ActivityData>();

            if (isTest)
            {
                WebOutlookCrmST.smmActivities.SMMACTIVITIESDataTable actDataTable = serviceSTG.GetsmmActivities(company, responsible);
                foreach (WebOutlookCrmST.smmActivities.SMMACTIVITIESRow row in actDataTable)
                {
                    ActivityData actData = new ActivityData(row);
                    actList.Add(actData);
                }
            }
            else
            {
                //company = "VAR";
                //responsible = "NRend";
                object[] results = serviceLIV.GetsmmActivities2(company, responsible);

                WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable actDataTable = new WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable();
                //actDataTable.DataSet.EnforceConstraints = false;
                //actDataTable = ((WebOutlookCrm.smmActivities.SMMACTIVITIESDataTable)(results[0]));
                try 
                {
                    actDataTable =
                        serviceLIV.GetsmmActivities(company, responsible);

                }
                catch (ConstraintException ex)
                {
                    DataRow[] rowErrors = actDataTable.GetErrors();
                    System.Diagnostics.Debug.WriteLine("smmActivities Errors:"
        + rowErrors.Length);

                    for (int i = 0; i < rowErrors.Length; i++)
                    {
                        System.Diagnostics.Debug.WriteLine(rowErrors[i].RowError);

                        foreach (DataColumn col in rowErrors[i].GetColumnsInError())
                        {
                            System.Diagnostics.Debug.WriteLine(col.ColumnName
                                + ":" + rowErrors[i].GetColumnError(col));
                        }
                    }
                }

                foreach (WebOutlookCrm.smmActivities.SMMACTIVITIESRow row in actDataTable)
                {
                    ActivityData actData = new ActivityData(row);
                    actList.Add(actData);
                }
            }

            return actList;
        }

        /*public void CheckDataSet(DataSet dataSet)
        {
            Assembly assembly = Assembly.LoadFrom(@"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Data.dll");
            Type type = assembly.GetType("System.Data.ConstraintEnumerator");
            ConstructorInfo ctor = type.GetConstructor(new[] { typeof(DataSet) });
            object instance = ctor.Invoke(new object[] { dataSet });
            BindingFlags bf = BindingFlags.Instance | BindingFlags.Public;
            MethodInfo m_GetNext = type.GetMethod("GetNext", bf);

            while ((bool)m_GetNext.Invoke(instance, null))
            {
                bool flag = false;
                MethodInfo m_GetConstraint = type.GetMethod("GetConstraint", bf);
                Constraint constraint = (Constraint)m_GetConstraint.Invoke(instance, null);
                Type constraintType = constraint.GetType();
                BindingFlags bfInternal = BindingFlags.Instance | BindingFlags.NonPublic;
                MethodInfo m_IsConstraintViolated = constraintType.GetMethod("IsConstraintViolated", bfInternal);
                flag = (bool)m_IsConstraintViolated.Invoke(constraint, null);
                if (flag)
                    Console.WriteLine("Constraint violated, ConstraintName: " + constraint.ConstraintName + ", tableName: " + constraint.Table);
            }

            foreach (DataTable table in dataSet.Tables)
            {
                foreach (DataColumn column in table.Columns)
                {
                    Type columnType = column.GetType();
                    BindingFlags bfInternal = BindingFlags.Instance | BindingFlags.NonPublic;

                    bool flag = false;
                    if (!column.AllowDBNull)
                    {
                        MethodInfo m_IsNotAllowDBNullViolated = columnType.GetMethod("IsNotAllowDBNullViolated", bfInternal);
                        flag = (bool)m_IsNotAllowDBNullViolated.Invoke(column, null);
                        if (flag)
                        {
                            Console.WriteLine("DBnull violated  --> ColumnName: " + column.ColumnName + ", tableName: " + column.Table.TableName);
                        }
                    }
                    if (column.MaxLength >= 0)
                    {
                        MethodInfo m_IsMaxLengthViolated = columnType.GetMethod("IsMaxLengthViolated", bfInternal);
                        flag = (bool)m_IsMaxLengthViolated.Invoke(column, null);
                        if (flag)
                            Console.WriteLine("MaxLength violated --> ColumnName: " + column.ColumnName + ", tableName: " + column.Table.TableName);
                    }
                }
            }
        }*/

        public bool DeleteActivity(string OUTLOOKENTRYID, string company)
        {
            return isTest ? serviceSTG.DeleteActivity(OUTLOOKENTRYID, company) : serviceLIV.DeleteActivity(OUTLOOKENTRYID, company);
        }

        public string InsertActivity(DateTime startDate, DateTime endDate, string emplId, bool allDay, string billingInformation, string body,
            string subject, string categories, int taskPriority, string location, string mileage, bool reminderSet, int reminderMinutesBeforeStart,
            string resources, bool responseRequested, int sensitivity, int activityTimeType, string domainUserId, string company,
            string OUTLOOKENTRYID)
        {
            return isTest ?
                serviceSTG.InsertActivity(startDate, endDate, emplId, allDay, billingInformation, body, subject, categories, taskPriority, location, mileage,
                reminderSet, reminderMinutesBeforeStart, resources, responseRequested, sensitivity, activityTimeType, domainUserId, company, OUTLOOKENTRYID) :

                serviceLIV.InsertActivity(startDate, endDate, emplId, allDay, billingInformation, body, subject, categories, taskPriority, location,
                mileage, reminderSet, reminderMinutesBeforeStart, resources, responseRequested, sensitivity, activityTimeType, domainUserId, company,
                OUTLOOKENTRYID);
        }

        public bool UpdateActivity(DateTime startDate, DateTime endDate, string emplId, bool allDay, string billingInformation, string body,
            string subject, string categories, int taskPriority, string location, string mileage, bool reminderSet, int reminderMinutesBeforeStart,
            string resources, bool responseRequested, int sensitivity, int activityTimeType, string domainUserId, string company,
            string OUTLOOKENTRYID, string activityNumber)
        {
            return isTest ?
                serviceSTG.UpdateActivity(startDate, endDate, emplId, allDay, billingInformation, body, subject, categories, taskPriority, location,
                mileage, reminderSet, reminderMinutesBeforeStart, resources, responseRequested, sensitivity, activityTimeType, domainUserId, company,
                OUTLOOKENTRYID, activityNumber) :

                serviceLIV.UpdateActivity(startDate, endDate, emplId, allDay, billingInformation, body, subject, categories, taskPriority, location, 
                mileage, reminderSet, reminderMinutesBeforeStart, resources, responseRequested, sensitivity, activityTimeType, domainUserId, company,
                OUTLOOKENTRYID, activityNumber);
        }

        public bool UpdateActivityOutlookEntryId(string activityNumber, string company, string OUTLOOKENTRYID)
        {
            return isTest ? serviceSTG.UpdateActivityOutlookEntryId(activityNumber, company, OUTLOOKENTRYID) :
                serviceLIV.UpdateActivityOutlookEntryId(activityNumber, company, OUTLOOKENTRYID);
        }

        public bool IsDeletedActivity(string OUTLOOKENTRYID, string company)
        {
            return isTest ? serviceSTG.IsDeletedActivity(OUTLOOKENTRYID, company) :
                serviceLIV.IsDeletedActivity(OUTLOOKENTRYID, company);
        }

        public List<BusRelData> GetBusinessRelationsBySalesDistrict(string domainUserId, string company, string salesDistrictId, string brNameFilter)
        {
            List<BusRelData> busRelList = new List<BusRelData>();

            if (isTest)
            {
                WebOutlookCrmST.Northwind.SMMBUSRELTABLE_DisplayDataTable table =
                    serviceSTG.GetBusinessRelationsBySalesDistrict(domainUserId, company, salesDistrictId, brNameFilter);

                foreach (WebOutlookCrmST.Northwind.SMMBUSRELTABLE_DisplayRow row in table)
                {
                    BusRelData busRelData = new BusRelData(row);
                    busRelList.Add(busRelData);
                }
            }
            else
            {
                WebOutlookCrm.Northwind.SMMBUSRELTABLE_DisplayDataTable table =
                    serviceLIV.GetBusinessRelationsBySalesDistrict(domainUserId, company, salesDistrictId, brNameFilter);

                foreach (WebOutlookCrm.Northwind.SMMBUSRELTABLE_DisplayRow row in table)
                {
                    BusRelData busRelData = new BusRelData(row);
                    busRelList.Add(busRelData);
                }
            }

            return busRelList;
        }

        public List<ContactPerson> GetContactsByPartyID(string partyId, string company)
        {
            List<ContactPerson> contactList = new List<ContactPerson>();

            if (isTest)
            {
                WebOutlookCrmST.Northwind.CONTACTPERSONShortDataTable table = serviceSTG.GetContactsByPartyID(partyId, company);

                foreach(WebOutlookCrmST.Northwind.CONTACTPERSONShortRow row in table)
                {
                    ContactPerson contactPerson = new ContactPerson(row);
                    contactList.Add(contactPerson);
                }
            }
            else
            {
                WebOutlookCrm.Northwind.CONTACTPERSONShortDataTable table = serviceLIV.GetContactsByPartyID(partyId, company);

                foreach (WebOutlookCrm.Northwind.CONTACTPERSONShortRow row in table)
                {
                    ContactPerson contactPerson = new ContactPerson(row);
                    contactList.Add(contactPerson);
                }
            }

            return contactList;
        }

        public List<SalesDistrict> GetUserSalesDistricts(string domainUserId, string company)
        {
            List<SalesDistrict> sdList = new List<SalesDistrict>();
            if (isTest)
            {
                WebOutlookCrmST.SMMTABLES.SalesDistrictDataTable table = serviceSTG.GetUserSalesDistricts(domainUserId, company);
                foreach (WebOutlookCrmST.SMMTABLES.SalesDistrictRow row in table)
                {
                    SalesDistrict salesDistrict = new SalesDistrict(row);
                    sdList.Add(salesDistrict);
                }
            }
            else
            {
                WebOutlookCrm.SMMTABLES.SalesDistrictDataTable table = serviceLIV.GetUserSalesDistricts(domainUserId, company);
                foreach (WebOutlookCrm.SMMTABLES.SalesDistrictRow row in table)
                {
                    SalesDistrict salesDistrict = new SalesDistrict(row);
                    sdList.Add(salesDistrict);
                }
            }

            return sdList;
        }
    }
}
