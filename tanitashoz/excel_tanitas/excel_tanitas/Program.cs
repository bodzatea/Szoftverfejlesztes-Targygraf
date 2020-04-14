using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace excel_tanitas
{
    class Program
    {
        _Application excel = new Application();
        Workbook wb;
        Worksheet ws;

        static void Main(string[] args)
        {
            Program p = new Program();
            p.readInFiles();
        }

        public void readInFiles()
        {
            var allFiles = Directory.EnumerateFiles(@"D:\Szoftverfejlesztes-Targygraf\tanitas_excelek");
            Regex fileFilter = new Regex(@"^(.)*?\\[^~$]+.xlsx$");
            foreach (string file in allFiles)
            {
                if (fileFilter.IsMatch(file)) //ha új excel fájlt találunk
                {
                    wb = excel.Workbooks.Open(file);

                    mergeSheets(); //az első sheetre rakunk mindent              
                    ws = wb.Worksheets[1]; //ws mindig az adott workbook első sheetje
                    writeOutData();

                    wb.Close(0); //bezárjuk, nem mentünk
                }
            }
            Console.ReadKey();
        }

        public void mergeSheets()
        { //minden sheetet átrakja a  legelsőre
            for (int i = 2; i < wb.Worksheets.Count+1; i++)
            {
                int goalSheetRow = getRowNum(wb.Worksheets[1]); //első sheet utolsó használt sora
                int sourceSheetRow = getRowNum(wb.Worksheets[i]); //i. sheet utolsó használt sora

                int sourceSheetColumn = getColNum(wb.Worksheets[i]); //i.sheet utolsó használt oszlopa

                string sourceSheetEnd = GetExcelColumnName(sourceSheetColumn) + sourceSheetRow;
                string goalSheetStart = "A" + (goalSheetRow + 1);

                wb.Worksheets[i].Range["A1", sourceSheetEnd].Copy(wb.Worksheets[1].Range[goalSheetStart]);
            }
            //wb.SaveCopyAs(@"D:\Szoftverfejlesztes-Targygraf\Book1.xlsx"); //testing, ha ki akarjuk menteni valamelyik sheetet
        }

        public void writeOutData()
        {
            int lastRow = getRowNum(ws);

            for (int currentRow = 1; currentRow < lastRow+1; currentRow++)
            {

                //if (ws.Cells[currentRow, 1].Value2 != null && ws.Cells[currentRow, 2].Value2 != null)
                //{
                    Console.WriteLine(ws.Cells[currentRow, 1].Value2+" "+ ws.Cells[currentRow, 2].Value2);
                string a= ws.Cells[currentRow, 1].Value2.ToString();
                //}
            }
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
