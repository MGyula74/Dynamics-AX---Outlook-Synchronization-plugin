using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrmOutlookAddIn2013.Model
{
    class Activity
    {
        public DateTime STARTDATETIME { get; set; }
        public DateTime ENDDATETIME { get; set; }
        public string ACTIVITYNUMBER { get; set; }
        public DateTime MODIFIEDDATETIME { get; set; }
        public string PURPOSE { get; set; }
        public int ALLDAY { get; set; }
        public string BILLINGINFORMATION { get; set; }
        public string USERMEMO { get; set; }
        public string OUTLOOKCATEGORIES { get; set; }
        public string LOCATION { get; set; }
        public string MILEAGE { get; set; }
        public int REMINDERACTIVE { get; set; }
        public int REMINDERMINUTES { get; set; }
        public string OUTLOOKRESOURCES { get; set; }
        public int RESPONSEREQUESTED { get; set; }
        public int SENSITIVITY { get; set; }
        public int ACTIVITYTIMETYPE { get; set; }

    }
}
