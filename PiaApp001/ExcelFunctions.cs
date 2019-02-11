using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace PiaApp001
{
    class ExcelFunctions
    {
        private static Excel.Application MyApp = null;

        private static Excel.Workbook MyBook = null;
        private static Excel.Workbook MyBookOut = null;

        private static Excel.Worksheet MySheet = null;
        private static Excel.Worksheet MySheetOut = null;

        private static Excel.Range MyRange = null;


        public void CreateExcel(string DB_PATH)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            List<string> EmpList = new List<string> { };



            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "F" + index.ToString()).Cells.Value;


                String Firstname = MyValues.GetValue(1, 4).ToString();
                String LastName = MyValues.GetValue(1, 5).ToString();
                EmpList.Add(Firstname + " " + LastName);


                MyRange = MySheet.Cells[index, 6].value(Firstname + " " + LastName);
            }
            MyBook.Save();
        }
        public void CreateExcel300Udd(string Inpath, string Outpath)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(Inpath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            List<string> Udd = new List<string> { };

            MyBookOut = MyApp.Workbooks.Open(Outpath);
            MySheetOut = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here


            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "F" + index.ToString()).Cells.Value;


                String UddNavn = MyValues.GetValue(1, 1).ToString();
                Udd.Add("TEC-udd:"+ UddNavn);


                MyRange = MySheet.Cells[index, 1].value("TEC-udd: "+UddNavn);
            }
            MyBook.Save();
        }
        public void CreateExcel300Navne(string Inpath, string Outpath)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(Inpath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            List<string> EmpList = new List<string> { };


            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "F" + index.ToString()).Cells.Value;


                String Firstname = MyValues.GetValue(1, 4).ToString();
                String LastName = MyValues.GetValue(1, 5).ToString();
                EmpList.Add(Firstname + " " + LastName);


                MyRange = MySheet.Cells[index, 6].value(Firstname + " " + LastName);
            }
            MyBook.Save();
        }
    }
}
