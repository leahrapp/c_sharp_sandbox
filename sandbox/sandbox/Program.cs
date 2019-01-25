using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Linq; 
using System.IO;
using System.Diagnostics;
using System.Web;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq.Expressions;
using System.Timers; 


using System.Configuration;

using System.Linq.Expressions;


namespace sandbox
{
    class Program
    {
        private static System.Timers.Timer aTimer;
        static void Main(string[] args)
        {



            //var path = AppDomain.CurrentDomain.BaseDirectory; // HttpContext.Current.Server.MapPath("~");

            //var msdsPath = @"C:\Users\lrapp\desktop\temp\666913_999000366076_999000365572_16JAN18 (4).pdf";

            //ItextExp(msdsPath);

           // Convert();
             SavePdf();


        }


        public static void Convert()
        {

            Console.WriteLine(DateTime.Now);


            var table = new DataTable();
            IWorkbook workbook;
            ISheet sheet = null; 
            using (FileStream FS = new FileStream(@"C:\temp\lowtech.xlsx", FileMode.Open, FileAccess.Read))
            {

                workbook = WorkbookFactory.Create(FS);
                sheet = workbook.GetSheetAt(0); 
               
                IRow headerRow = sheet.GetRow(0); 
                foreach (ICell headerCell in headerRow)
                {
                    var hCell = headerCell.ToString();
                    hCell.Replace('/', '_').Replace(' ', '_').Replace("#", "Num");
                    table.Columns.Add(hCell); 
                }
                int rowIndex = 0;
                foreach (IRow row in sheet)
                {
                    // skip header row
                    if (rowIndex++ == 0) continue;
                    DataRow dataRow = table.NewRow();
                    dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
                    table.Rows.Add(dataRow);
                }

            }
            var bleh = table; 
            Console.WriteLine(DateTime.Now);
            Console.ReadLine(); 


        }


        public static void SavePdf()
        {
            Console.WriteLine(DateTime.Now);
            //var path = HttpContext.Current.Server.MapPath("~");
            //var msdsPath = string.Format("{0}\\resources\\MSDS", path);
            var manualMsdsPath = string.Format("C:\\Users\\lrapp\\Desktop\\LowTechApp\\LowTechApp");
            var msdsDirectory = Directory.EnumerateFiles(@"C:\Users\lrapp\Desktop\LowTechApp\LowTechApp").ToList();

          

           
            var connSql = @"Data Source=DESKTOP-3Q0952R;Initial Catalog=sandbox;Integrated Security=True;";

            foreach (var m in msdsDirectory)
            {
                if (m.Contains(".PDF"))
                {
                    var queryString = string.Format(" insert into pdftester(pdf) select [Doc_Data].* from openrowset (bulk '{0}', single_blob) [Doc_Data]",m);

                    using (SqlConnection connection = new SqlConnection(connSql))
                    {
                        try
                        {
                            SqlCommand command = new SqlCommand(queryString, connection);

                            command.Connection.Open();
                            command.ExecuteNonQuery();
                        }
                        catch
                        {



                        }


                    }
                }
            }
            Console.WriteLine(DateTime.Now);
            Console.ReadLine();

        }
        public static void ItextExp(string path)
        {
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
            var count = 0;
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string thePage = PdfTextExtractor.GetTextFromPage(reader, i, its);
                    string[] theLines = thePage.Split('\n');
                    foreach (var theLine in theLines)
                    {
                        text.AppendLine(theLine);
                    
                    }
                }

                
                Console.WriteLine(text.ToString()); 
                Console.ReadLine();
            }



        }
        public void foo()
        {
            var path = AppDomain.CurrentDomain.BaseDirectory; // HttpContext.Current.Server.MapPath("~");
            var msdsPath = string.Format("{0}\\resources\\MSDS", path);


            var manualMsdsPath = string.Format("{0}\\resources\\manualMSDS", path).GetHashCode();

            var msdsDirectory = Directory.EnumerateFiles(@"C:\source\repos\HAZMAT.lrapp\resources\MSDS").Count();
            // var count = msdsDirectory.

            var manualMsdsDirectory = Directory.EnumerateFiles(@"C:\source\repos\HAZMAT.lrapp\resources\manualMSDS").GetHashCode();
            Console.WriteLine("Manual: " + manualMsdsDirectory);
            Console.WriteLine("MSDS: " + msdsDirectory);
            Console.ReadLine();
        }

    }
}
