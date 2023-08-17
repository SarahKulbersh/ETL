using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace TRY
{
    internal class Program
    {


        static void ExtractAndTransformData()
        {
            String filePath = "C:\\Users\\Owner\\Desktop\\employee3.xlsx";

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];


            //These two lines do the magic.
            ws.Columns.ClearFormats();
            ws.Rows.ClearFormats();

            int iTotalColumns = ws.UsedRange.Columns.Count;
            int iTotalRows = ws.UsedRange.Rows.Count;


            List<Employee> employees = new List<Employee>();
            
            for (int i = 0; i < iTotalRows; i++)
            {
                double id = (double)ws.Cells[i + 1, 1].Value;
                string firstName = (string)ws.Cells[i + 1, 2].Value;
                string lastName = (string)ws.Cells[i + 1, 3].Value;
                double age = (double)ws.Cells[i + 1, 4].Value;
                Employee employee = new Employee(id, firstName, lastName, age);
                Console.WriteLine(employee.FirstName);
                employees.Add(employee);
            }

             employees.OrderBy(o => o.Age).ToList();


            string path ="";
            string data;
            string[] rowsArray = new string[employees.Count];

            Console.WriteLine(employees.Count);

            for (int i = 0;i < employees.Count;i++)
            {
                Console.WriteLine("check"+employees[i].FirstName);
                data = $"{employees[i].Id},{employees[i].FirstName},{employees[i].LastName},{employees[i].Age}";
                rowsArray[i] = data;
                File.WriteAllLines("C:\\Users\\Owner\\Desktop\\employee99.csv", rowsArray);
            }


            Console.WriteLine(iTotalRows);
            Console.ReadLine();

        }



        static void Main(string[] args)
        {

            Console.WriteLine("Hello World!");
            ExtractAndTransformData();

        }

        sealed class ETLProcessor
        {
            private ETLProcessor() { }
            private static ETLProcessor instance = null;
            public static ETLProcessor Instance
            {
                get
                {
                    if (instance == null)
                    {
                        instance = new ETLProcessor();
                    }
                    return instance;
                }
            }
        }


        private enum DataSource
        {
            CSV,
            database,
            API
        }
    }
}
