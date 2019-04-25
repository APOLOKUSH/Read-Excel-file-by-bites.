using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace EA_3_
{


    class Program
    {
      
        List<List<string>> excelData = new List<List<string>>();
        List<List<string>> val = new List<List<string>>();
        List<string> v1 = new List<string>();
        List<string> v2 = new List<string>();
        List<string> v3 = new List<string>();


        static public void Main( string[] args )  //string[] args
        {
            //create a list to hold all the values
            // List<List<string>> excelData = new List<List<string>>();
            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes("C:\\Users\\mykolakandiuk\\Documents\\totalbetaEurope.xls");

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))  
            {
                //loop all worksheets
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <=
                       worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {


                            }
                        }
                    }
                }
            }

            


        }
        
    }
}

