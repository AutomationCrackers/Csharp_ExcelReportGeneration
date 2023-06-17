using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Csharp_ExcelReportGeneration
{
  public static  class ObjectMapper
    {
        public static List<Model> MapDataTableToModel(string input, DataTable table)
        {
            List<Model> datacolumn = new List<Model>();
            Model dt;
            Console.WriteLine("Mapping Datatable to Objects");
            for(int row = 0; row <table.Rows.Count; row++)
            {
                int i = 0;
                dt = new Model();
                dt.Environments = table.Rows[row][i].ToString();
                i++;
                dt.Status = table.Rows[row][i].ToString();
                i++;
                dt.URL = table.Rows[row][i].ToString();
                i++;
                dt.UserId = table.Rows[row][i].ToString();
                i++;
                dt.Password = table.Rows[row][i].ToString();
                i++;
                dt.LoginResult = table.Rows[row][i].ToString();
                i++;
                datacolumn.Add(dt);
            }
            Console.WriteLine("Mapping Datatable to Objects - Completed");
            return datacolumn;
        }
    }
}
