using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Csharp_ExcelReportGeneration
{
  public static  class ExcelReportGenerator
    {
        public static string excelReportPath, environments, reportname;
        public static XLWorkbook wb;
        public static IXLWorksheet Ws, Ws1;

        //Step 1 Define Excel Report Path
        public static void GetExcelReportPath()
        {
            string currentDirectoryPath = Environment.CurrentDirectory;
            string actualPath = currentDirectoryPath.Substring(0, currentDirectoryPath.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            excelReportPath = projectPath + "\\ExcelReports\\";
        }

        //Step 2 - Delete the Old Exisitng Reports inside the Path
        public static void DeleteOldExcelReports()
        {
            GetExcelReportPath();
            if (Directory.Exists(excelReportPath))
            {
                string[] ExcelReports = Directory.GetFiles(excelReportPath);
                foreach (string Reports in ExcelReports)
                {
                    File.Delete(excelReportPath);
                }
            }
        }
        //Step 3 - Define the Excel Report name
        public static void ConfigureExcelReport()
        {
            DeleteOldExcelReports();
            GetExcelReportPath();
            reportname = excelReportPath + "\\ExcelReports\\AutomationTestReport"+DateTime.Now.ToString("d-M-yyyy")+DateTime.Now.ToString("hhmmss tt")+".XLSX";
        }

        //Step4 - Excel Styling - Boder, theme color, font..etc
        public static void EConfigurexcelStyling()
        {
            wb = new XLWorkbook();
            
            //Sheet Names
            Ws = wb.Worksheets.Add("TestResults");
            Ws = wb.Worksheets.Add("Logs");

            Ws.Range("B3:P11").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Ws.Range("C4:O4").Style.Border.OutsideBorder = XLBorderStyleValues.None;
            Ws.Range("C4:O4").Style.Border.OutsideBorder = XLBorderStyleValues.None;
            Ws.Range("C5:O9").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Ws.Range("C5:O9").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            Ws.Range("B3:P3").Merge().Style.Fill.BackgroundColor = XLColor.Blue;
            //Report Header
            Ws.Row(3).Cell("B").Value = "Automation Test Report";
            Ws.Row(3).Cell("B").Style.Font.FontColor = XLColor.White;
            Ws.Row(3).Cell("B").Style.Font.SetFontSize(18);
            Ws.Row(3).Cell("B").Style.Font.SetBold(true);
            Ws.Row(3).Cell("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            //Log the Date when is the report generated
            Ws.Row(5).Cell("B").Value = "Date: " + DateTime.Now.ToString("dd/MM/yyyy");
            Ws.Row(5).Cell("B").Style.Font.SetBold(true);

            Ws.Range("D5:O5").Merge().Style.Fill.BackgroundColor = XLColor.LightBlue;
            Ws.Row(5).Cell("D").Value = "Status @ " + DateTime.Now.ToString("hh:mm tt", CultureInfo.InvariantCulture) + "BST";
            Ws.Row(5).Cell("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            Ws.Row(5).Cell("D").Style.Font.SetBold(true);
           
            //Validation Row headers
            Ws.Row(6).Cell("3").Value = "Environments";
            Ws.Row(6).Cell("3").Style.Font.SetBold(true);

            Ws.Row(7).Cell("3").Value = "LoginResult";
            Ws.Row(7).Cell("3").Style.Font.SetBold(true);

        }

        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            foreach(PropertyInfo prop in props)
            {
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>));
                dataTable.Columns.Add(prop.Name);
            }
            foreach(T Items in items)
            {
                var values = new object[props.Length];
                for(int i=0; i<props.Length; i++)
                {
                    values[i] = props[i].GetValue(items, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        //Step5 -Log Result to Excel
        public static void LogResultToExcel(List<Model> model, string Environment)
        {
            DataTable dt = new DataTable();
            dt = ToDataTable(model);

            int i = 4;
            int col = 67; // Based on how many environments for example i took 11 environments- cell number on the 11th env is 67
            foreach(Model md in model)
            {
                char c = (char)col;
                if(md.logs.Rows.Count>0)
                {
                    if(Ws1.RowsUsed().Count()==0)
                    {
                        IXLCell CellForNewData = Ws1.Cell(1, 1);
                        CellForNewData.InsertTable(md.logs);
                    }
                    else
                    {
                        int RowNumber = Ws1.LastRowUsed().RowNumber();
                        IXLCell CellForNewData = Ws1.Cell(RowNumber+1, 1);
                        CellForNewData.InsertTable(md.logs);
                    }
                }
                environments = md.Environments;
                Ws.Row(6).Cell("i").Value = md.Environments;
                Ws.Row(6).Cell("i").Style.Font.FontColor = XLColor.Black;
                Ws.Row(6).Cell("i").Style.Font.SetBold(true);
              

                if(environments.Equals("SIT"))
                {
                    Ws.Row(7).Cell("i").Value = md.LoginResult;
                    Ws.Row(7).Cell("i").Style.Font.FontColor = md.LoginResultColor;
                }
                i = i + 1;
                col++;
            }

            
        }
    }
}
