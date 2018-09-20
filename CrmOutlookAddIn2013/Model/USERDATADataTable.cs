using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace CrmOutlookAddIn2013.Model
{
    public class USERDATADataTable : DataTable
    {
        public USERDATADataRow this[int index]
        {
            get { return (USERDATADataRow)Rows[index]; }
        }
    }

    public class USERDATADataRow : DataRow
    {
        internal USERDATADataRow(DataRowBuilder builder): base(builder) {
        }

        public string OUTLOOKCALENDAROUTLOOKENTRYID
        {
            get { return (string)base["OUTLOOKCALENDAROUTLOOKENTRYID"]; }
            set { base["OUTLOOKCALENDAROUTLOOKENTRYID"] = value; }
        }
    }
}
