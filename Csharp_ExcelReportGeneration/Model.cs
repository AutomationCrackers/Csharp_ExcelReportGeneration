using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Csharp_ExcelReportGeneration
{
    public class Model
    {
        public DataTable logs = new DataTable();

        //Constructor
        public Model()
        {
            //how many columns we want is configured -  I configured as 3 columns
            logs.Columns.AddRange(new DataColumn[3]
                {
                    //name of the columns
                    new DataColumn("Date"),
                    new DataColumn("Method Name"),
                    new DataColumn("Log Message"),
                });
        }

            //Variables
        public string Environments { get; set; }
        public string Status { get; set; }
        public string URL { get; set; }
        public string UserId { get; set; }
        public string Password { get; set; }
        public string LoginResult { get; set; }

        public XLColor LoginResultColor = XLColor.Black;
    }
}
