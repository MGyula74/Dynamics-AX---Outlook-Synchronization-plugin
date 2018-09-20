using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmOutlookAddIn2013.Model
{
    class UserData
    {
        public string networkDomain { get; set; }
        public string networkAlias { get; set; }
        public string OUTLOOKCALENDAROUTLOOKENTRYID { get; set; }
        public string OUTLOOKCALENDAROUTLOOKSTOREID { get; set; }
        public string COMPANY { get; set; }
        public string EMPLID { get; set; }
        public int smmSynchronizeDaysBack { get; set; }
        public int smmSynchronizeDaysForward { get; set; }
        public string CompanyCountryId { get; set; }
    }

    public class ActivityData
    {
        public string ACTIVITYNUMBER { get; set; }
        public DateTime STARTDATETIME { get; set; }
        public DateTime ENDDATETIME { get; set; }
        public string OUTLOOKENTRYID { get; set; }
        public DateTime MODIFIEDDATETIME { get; set; }
        public string PURPOSE { get; set; }
        public int ALLDAY { get; set; }
        public string BILLINGINFORMATION { get; set; }
        public string USERMEMO { get; set; }
        public string OUTLOOKCATEGORIES { get; set; }
        public int TASKPRIORITY { get; set; }
        public string LOCATION { get; set; }
        public string MILEAGE { get; set; }
        public int REMINDERACTIVE { get; set; }
        public int REMINDERMINUTES { get; set; }
        public string OUTLOOKRESOURCES { get; set; }
        public int RESPONSEREQUESTED { get; set; }
        public int SENSITIVITY { get; set; }
        public int ACTIVITYTIMETYPE { get; set; }

        public string CUSTNAME { get; set; }
        public string CUSTABC { get; set; }
        public string CUSTSALESREP { get; set; }
        public string DLVMODE { get; set; }
        public string NAMEALIAS { get; set; }
        public string CONTACTPERSONID { get; set; }
        public string CONTACTNAME { get; set; }
        public string CONTACTPHONE { get; set; }
        public string CONTACTEMAIL { get; set; }

        public ActivityData(WebOutlookCrmST.smmActivities.SMMACTIVITIESRow activityRow)
        {
            this.ACTIVITYNUMBER = activityRow.ACTIVITYNUMBER;
            this.STARTDATETIME = activityRow.STARTDATETIME;
            this.ENDDATETIME = activityRow.ENDDATETIME;
            this.OUTLOOKENTRYID = activityRow.OUTLOOKENTRYID;
            this.MODIFIEDDATETIME = activityRow.MODIFIEDDATETIME;
            this.PURPOSE = activityRow.PURPOSE;
            this.ALLDAY = activityRow.ALLDAY;
            this.BILLINGINFORMATION = activityRow.BILLINGINFORMATION;
            this.USERMEMO = activityRow.USERMEMO;
            this.OUTLOOKCATEGORIES = activityRow.OUTLOOKCATEGORIES;
            this.TASKPRIORITY = activityRow.TASKPRIORITY;
            this.LOCATION = activityRow.LOCATION;
            this.MILEAGE = activityRow.MILEAGE;
            this.REMINDERACTIVE = activityRow.REMINDERACTIVE;
            this.REMINDERMINUTES = activityRow.REMINDERMINUTES;
            this.OUTLOOKRESOURCES = activityRow.OUTLOOKRESOURCES;
            this.RESPONSEREQUESTED = activityRow.RESPONSEREQUESTED;
            this.SENSITIVITY = activityRow.SENSITIVITY;
            this.ACTIVITYTIMETYPE = activityRow.ACTIVITYTIMETYPE;
            this.CUSTNAME = activityRow.CUSTNAME;
            this.CUSTABC = activityRow.CUSTABC;
            this.CUSTSALESREP = activityRow.CUSTSALESREP;
            this.DLVMODE = activityRow.DLVMODE;
            this.NAMEALIAS = activityRow.NAMEALIAS;
            this.CONTACTPERSONID = activityRow.CONTACTPERSONID;
            this.CONTACTNAME = activityRow.CONTACTNAME;
            this.CONTACTPHONE = activityRow.CONTACTPHONE;
            this.CONTACTEMAIL = activityRow.CONTACTEMAIL;
        }

        public ActivityData(WebOutlookCrm.smmActivities.SMMACTIVITIESRow activityRow)
        {
            this.ACTIVITYNUMBER = activityRow.ACTIVITYNUMBER;
            this.STARTDATETIME = activityRow.STARTDATETIME;
            this.ENDDATETIME = activityRow.ENDDATETIME;
            this.OUTLOOKENTRYID = activityRow.OUTLOOKENTRYID;
            this.MODIFIEDDATETIME = activityRow.MODIFIEDDATETIME;
            this.PURPOSE = activityRow.PURPOSE;
            this.ALLDAY = activityRow.ALLDAY;
            this.BILLINGINFORMATION = activityRow.BILLINGINFORMATION;
            this.USERMEMO = activityRow.USERMEMO;
            this.OUTLOOKCATEGORIES = activityRow.OUTLOOKCATEGORIES;
            this.TASKPRIORITY = activityRow.TASKPRIORITY;
            this.LOCATION = activityRow.LOCATION;
            this.MILEAGE = activityRow.MILEAGE;
            this.REMINDERACTIVE = activityRow.REMINDERACTIVE;
            this.REMINDERMINUTES = activityRow.REMINDERMINUTES;
            this.OUTLOOKRESOURCES = activityRow.OUTLOOKRESOURCES;
            this.RESPONSEREQUESTED = activityRow.RESPONSEREQUESTED;
            this.SENSITIVITY = activityRow.SENSITIVITY;
            this.ACTIVITYTIMETYPE = activityRow.ACTIVITYTIMETYPE;
            this.CUSTNAME = activityRow.CUSTNAME;
            this.CUSTABC = activityRow.CUSTABC;
            this.CUSTSALESREP = activityRow.CUSTSALESREP;
            this.DLVMODE = activityRow.DLVMODE;
            this.NAMEALIAS = activityRow.NAMEALIAS;
            this.CONTACTPERSONID = activityRow.CONTACTPERSONID;
            this.CONTACTNAME = activityRow.CONTACTNAME;
            this.CONTACTPHONE = activityRow.CONTACTPHONE;
            this.CONTACTEMAIL = activityRow.CONTACTEMAIL;
        }
    }

    public class BusRelData
    {
        public string BUSRELACCOUNT { get; set; }
        public string CUSTACCOUNT { get; set; }
        public string FULLNAME { get; set; }
        public string PARTYID { get; set; }
        public int RELATIONTYPE { get; set; }

        public BusRelData(WebOutlookCrmST.Northwind.SMMBUSRELTABLE_DisplayRow row)
        {
            this.BUSRELACCOUNT = row.BUSRELACCOUNT;
            this.CUSTACCOUNT = row.IsNull("CUSTACCOUNT") ? String.Empty : row.CUSTACCOUNT;
            this.FULLNAME = row.FULLNAME;
            this.PARTYID = row.PARTYID;
            this.RELATIONTYPE = row.RELATIONTYPE;
        }

        public BusRelData(WebOutlookCrm.Northwind.SMMBUSRELTABLE_DisplayRow row)
        {
            this.BUSRELACCOUNT = row.BUSRELACCOUNT;
            this.CUSTACCOUNT = row.IsNull("CUSTACCOUNT") ? String.Empty : row.CUSTACCOUNT;
            this.FULLNAME = row.FULLNAME;
            this.PARTYID = row.PARTYID;
            this.RELATIONTYPE = row.RELATIONTYPE;
        }
    }

    public class ContactPerson
    {
        public string ADDRESS { get; set; }
        public string CONTACTPERSONID { get; set; }
        public string EMAIL { get; set; }
        public string NAME { get; set; }
        public string PARTYID { get; set; }
        public string PHONE { get; set; }

        public ContactPerson(WebOutlookCrmST.Northwind.CONTACTPERSONShortRow row)
        {
            this.ADDRESS = row.ADDRESS;
            this.CONTACTPERSONID = row.CONTACTPERSONID;
            this.EMAIL = row.EMAIL;
            this.NAME = row.NAME;
            this.PARTYID = row.PARTYID;
            this.PHONE = row.PHONE;
        }

        public ContactPerson(WebOutlookCrm.Northwind.CONTACTPERSONShortRow row)
        {
            this.ADDRESS = row.ADDRESS;
            this.CONTACTPERSONID = row.CONTACTPERSONID;
            this.EMAIL = row.EMAIL;
            this.NAME = row.NAME;
            this.PARTYID = row.PARTYID;
            this.PHONE = row.PHONE;
        }
    }

    public class SalesDistrict
    {
        public string DESCRIPTION { get; set; }
        public string SALESDISTRICTID { get; set; }
        public string SMMBUSRELACCRESPONSIBLE { get; set; }

        public SalesDistrict(WebOutlookCrmST.SMMTABLES.SalesDistrictRow row)
        {
            this.DESCRIPTION = row.DESCRIPTION;
            this.SALESDISTRICTID = row.SALESDISTRICTID;
            this.SMMBUSRELACCRESPONSIBLE = row.SMMBUSRELACCRESPONSIBLE;
        }

        public SalesDistrict(WebOutlookCrm.SMMTABLES.SalesDistrictRow row)
        {
            this.DESCRIPTION = row.DESCRIPTION;
            this.SALESDISTRICTID = row.SALESDISTRICTID;
            this.SMMBUSRELACCRESPONSIBLE = row.SMMBUSRELACCRESPONSIBLE;
        }
    }

}
