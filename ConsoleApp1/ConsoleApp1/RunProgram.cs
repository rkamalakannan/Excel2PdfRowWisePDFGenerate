using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConsoleApp1
{
    class RunProgram
    {

        public static void Main()

        {

            string path = @"C:\File\"; //Paste your excel in this folder
            Workbook wb = new Workbook(path + "Book1.xlsx");
            Worksheet worksheet = wb.Worksheets[0];

            foreach (var row in worksheet.Cells.Rows)
            {
                Workbook wb1 = new Workbook();
                Row row1 = row as Row;
                wb1.Worksheets[0].Cells.CopyRow(worksheet.Cells, row1.Index, 0);
                wb1.Save(path + $"{row1.Index}.pdf", SaveFormat.Pdf);
            }
        }
    }
}


