using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
namespace Targygraf
{
    class files
    {
        _Application excel = new Application();
        Workbook wb;
        Worksheet ws;

        public files() {
            //wb = excel.Workbooks.Open(@"d:\szoftverfejlesztes\proginfo.xlsx");
            // ws = wb.Worksheets[Sheet];

            var allFiles = Directory.EnumerateFiles(@"d:\szoftverfejlesztes\excelek");
            Regex fileFilter = new Regex(@"^(.)*?\\[^~$]+.xlsx$");
            List<string> excelFiles = new List<string>();
            foreach (string file in allFiles)
            {
                if (fileFilter.IsMatch(file))
                {
                    excelFiles.Add(file);
                    wb = excel.Workbooks.Open(file);

                    mergeSheets();
                    ws = wb.Worksheets[1];
                    readInData();

                    Console.WriteLine("");
                    wb.Close(0);
                }
            }

        }

        ~files() {
            
        }

        private string GetExcelColumnName(int columnNumber)
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

        public int getRowNum(Worksheet sheet) {
            int result=sheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, 
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            return result;
        }

        public int getColNum(Worksheet sheet)
        {
            int result=sheet.Cells.Find("*", System.Reflection.Missing.Value, 
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByColumns, 
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            return result;
        }

        public void mergeSheets() {

            for (int i = 2; i < wb.Worksheets.Count; i++) {
                int goalSheetRow = getRowNum(wb.Worksheets[1]);
                int sourceSheetRow = getRowNum(wb.Worksheets[i]);
                string twentyEnd = GetExcelColumnName(20) + sourceSheetRow;
                //merging sheets until the 20nd column
                wb.Worksheets[i].Range["A1", twentyEnd].UnMerge();

                int sourceSheetColumn = getColNum(wb.Worksheets[i]);
                
                string sourceSheetEnd = GetExcelColumnName(sourceSheetColumn+1) + sourceSheetRow;
                string goalSheetStart = "A" + (goalSheetRow+1);

                //Console.WriteLine(goalSheetStart+" "+sourceSheetRow+" "+sourceSheetColumn); //testing
                Console.WriteLine(goalSheetStart); //testing
                wb.Worksheets[i].Range["A1", sourceSheetEnd].Copy(wb.Worksheets[1].Range[goalSheetStart]);
            }

            //wb.SaveCopyAs(@"D:\szoftverfejlesztes\Book1.xlsx"); //testing

            
        }

        public string hunOnly(string text) {
            string result;
            MatchCollection matches = Regex.Matches(text, @"( )");
            int count = matches.Count / 2+1;
            MatchCollection matches2 = Regex.Matches(text, @"(.*?\s){"+count+"}");
            result= matches2[0].Groups[0].Value;
            return result;
        }


        public void readInSzak() {
            string szakNev = "";
            string szakTipus = "";
            string ervenyesseg = "";
         
            Regex nevFilter = new Regex(@"^(.*?)(BSc)|(MSc)");
            if (nevFilter.IsMatch(ws.Cells[1, 1].Value2.ToString()))
            {
                MatchCollection matches = Regex.Matches(ws.Cells[1, 1].Value2.ToString(), @"^(.*?)(BSc)|(MSc)");
                szakNev = matches[0].Groups[1].Value;
                szakTipus = matches[0].Groups[2].Value;
                Console.WriteLine(szakNev + " " + szakTipus); //testing
            }

            Regex ervenyessegFilter = new Regex(@"\w.*(\d{4}\/\d{2}.*){2}");
            if (ervenyessegFilter.IsMatch(ws.Cells[3, 1].Value2.ToString()))
            {
                MatchCollection matches = Regex.Matches(ws.Cells[3, 1].Value2.ToString(), @"\w.*(\d{4}\/\d{2}.*){2}");
                ervenyesseg = matches[0].Groups[0].Value;
                Console.WriteLine(ervenyesseg); //testing
            }
        }

        public int getKreditNum(int i) {
            int sourceSheetColumn = getColNum(ws);
            //Console.WriteLine(sourceSheetColumn); //testing
            int kreditColumn = 1;
            int notEmptyCells = 0;
            while ((kreditColumn < sourceSheetColumn) && (notEmptyCells < 4))
            {
                if ((ws.Cells[(i + 1), kreditColumn].Value2) != null)
                {
                    notEmptyCells++;
                }
                kreditColumn++;
            }
            kreditColumn--;
            //Console.WriteLine(kreditColumn); //testing

            return kreditColumn;
        }

        public int getUjFelev(int wsRowCount, int j, ref int ajanlottFelev, ref string targyTipus) {          
            int kreditColumn = getKreditNum(j);
            j += 2;
            Regex osszesFilter = new Regex(@"^Összesen$|^Összesítés$|^Kreditpontok a modelltanterv féléveiben nem tanári szakirányon$");

            while (j < wsRowCount && ws.Cells[j, 1].Value2 != null && !osszesFilter.IsMatch(ws.Cells[j, 1].Value2.ToString())) //amig nincs osszesen
            {
                if (ws.Cells[j, kreditColumn].Value2 == null) //ha uj targytipus talalunk
                {
                    ajanlottFelev = 0;
                    targyTipus = ws.Cells[j, 1].Value2.ToString();
                }
                else //ha uj tantargyot talalunk
                {
                    getUjTantargy(kreditColumn, j, ajanlottFelev, targyTipus);
                }

                j++;       
                
            }
            return j;

        }

       public void getUjTantargy(int kreditColumn, int j, int ajanlottFelev, string targyTipus) {
            string targyNev = "";
            string targyKod = "";
            int kredit = 0;

            string elofeltetelKod = "";
            int egyszerrefelveheto = 0;

            targyNev = Regex.Replace(ws.Cells[j, 1].Value2.ToString(), @"\r\n?|\n", " ");
            if (ws.Cells[j, 2].Value2 != null)
            {
                targyKod = ws.Cells[j, 2].Value2.ToString();
            }
            else targyKod = "";
            Console.WriteLine(targyNev + " " + targyKod);
            //targynev és targykod megvan, hozzaadas az adatbazishoz
            //ellenorzes hogy eddig nem volt ilyen

            string trim = ws.Cells[j, kreditColumn].Value2.ToString().Substring(0, 1);
            kredit = Convert.ToInt32(trim);
            Console.WriteLine(ajanlottFelev + " " + targyTipus + " " + kredit);
            //ajanlottfelev, targytipus, kredit
        }

        public void readInData()
        {
            SqliteDataAccess sql = new SqliteDataAccess();

            readInSzak();

            int ajanlottFelev = 1;
            string targyTipus = "Kötelező szakmai tárgyak";

            int wsRowCount = getRowNum(ws);
            
            
            for (int i = 1; i < wsRowCount; i++)
            {
                Regex felevFilter = new Regex(ajanlottFelev+". ?félév$|(?i)^Differenciált szakmai ismeretek$");
                if (ws.Cells[i, 1].Value2 != null && felevFilter.IsMatch(ws.Cells[i, 1].Value2.ToString())) //ha uj felev van
                {
                    i =getUjFelev(wsRowCount, i, ref ajanlottFelev, ref targyTipus);
                    

                    ajanlottFelev++;
                }
            }
        }
    }
}
  