using System;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace tanitas
{
    class Program
    {
        _Application excel = new Application();
        Workbook wb;
        Worksheet ws;

        void Main(string[] args)
        {
            readInFiles();
        }

        public void readInFiles() {
            var allFiles = Directory.EnumerateFiles(@"D:\Szoftverfejlesztes-Targygraf\tanitas_excelek");
            Regex fileFilter = new Regex(@"^(.)*?\\[^~$]+.xlsx$");
            foreach (string file in allFiles)
            {
                if (fileFilter.IsMatch(file)) //ha új excel fájlt találunk
                {
                    wb = excel.Workbooks.Open(file);

                    mergeSheets(); //az első sheetre rakunk mindent              
                    ws = wb.Worksheets[1]; //ws mindig az adott workbook első sheetje

                    wb.Close(0); //bezárjuk, nem mentünk
                }
            }
        }

        public void mergeSheets()
        { //minden sheetet átrakja a  legelsőre
            for (int i = 2; i < wb.Worksheets.Count; i++)
            {
                int goalSheetRow = getRowNum(wb.Worksheets[1]); //első sheet utolsó használt sora
                int sourceSheetRow = getRowNum(wb.Worksheets[i]); //i. sheet utolsó használt sora
                string twentyEnd = GetExcelColumnName(15) + sourceSheetRow; //meddig akarjuk hogy unmergeltek legyenek a cellák?

                int sourceSheetColumn = getColNum(wb.Worksheets[i]); //i.sheet utolsó használt oszlopa

                string sourceSheetEnd = GetExcelColumnName(sourceSheetColumn + 1) + sourceSheetRow;
                string goalSheetStart = "A" + (goalSheetRow + 1);

                //Console.WriteLine(goalSheetStart+" "+sourceSheetRow+" "+sourceSheetColumn); //testing
                //Console.WriteLine(goalSheetStart); //testing
                wb.Worksheets[i].Range["A1", sourceSheetEnd].Copy(wb.Worksheets[1].Range[goalSheetStart]);
            }
            //wb.SaveCopyAs(@"D:\Szoftverfejlesztes-Targygraf\Book1.xlsx"); //testing, ha ki akarjuk menteni valamelyik sheetet
        }

        public int getRowNum(Worksheet sheet) //visszadja a legutolsó nem üres sor sorszámát
        {
            int result = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows,
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            return result;
        }

        public int getColNum(Worksheet sheet) //visszadja a legutolsó nem üres oszlop sorszámát
        {
            int result = sheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByColumns,
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            return result;
        }

        private string GetExcelColumnName(int columnNumber) //szám alapján visszaadja a megfelelő betűt
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }




    }
}
}
