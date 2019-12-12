using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelOperation
{
    public class HelperMethod
    {
        public double sumTwoNumber(double firstNumber, double secondNumber)
        {
            return firstNumber + secondNumber;
        }

        public void mergeExcels(string inputFolderPath, string outputFolderPath, string outPutFileName)
        {
            string[] filePaths = Directory.GetFiles(inputFolderPath, "*.xlsx");
            if (filePaths.Length > 0)
            {
               // createExcelSheet();
                for (int i = 0; i < filePaths.Length; i++)
                {
                    Console.WriteLine(filePaths[i]);
                }
            }
          
        }

        private void createExcelSheet()
        {
            Application excel;
            Workbook worKbooK;
            Worksheet worKsheeT;
            Range celLrangE;
            excel = new Microsoft.Office.Interop.Excel.Application();
        }
    }
}
