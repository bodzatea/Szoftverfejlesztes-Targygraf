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
    class excelFiles
    {
        _Application excel = new Application();
        Workbook wb; //jelenlegi munkafüzet
        Worksheet ws; //az első sheetre mentünk át mindent

        public excelFiles() {
            //wb = excel.Workbooks.Open(@"d:\szoftverfejlesztes\proginfo.xlsx");
            //ws = wb.Worksheets[Sheet];

            var allFiles = Directory.EnumerateFiles(@"D:\Szoftverfejlesztes-Targygraf\excelek");
            Regex fileFilter = new Regex(@"^(.)*?\\[^~$]+.xlsx$");
            foreach (string file in allFiles)
            {
                if (fileFilter.IsMatch(file)) //ha új excel fájlt találunk
                {
                    wb = excel.Workbooks.Open(file);

                    mergeSheets(); //az első sheetre rakunk mindent
                    ws = wb.Worksheets[1]; //ws mindig az adott workbook első sheetje
                    readInData(); //beolvasunk

                    Console.WriteLine(""); //testing
                    wb.Close(0); //bezárjuk, nem mentünk
                }
            }
        }

        ~excelFiles() {
            
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

        public int getRowNum(Worksheet sheet) //visszadja a legutolsó nem üres sor sorszámát
        { 
            int result=sheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, 
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            return result;
        }

        public int getColNum(Worksheet sheet) //visszadja a legutolsó nem üres oszlop sorszámát
        {
            int result=sheet.Cells.Find("*", System.Reflection.Missing.Value, 
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByColumns, 
                XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            return result;
        }

        public void mergeSheets() { //minden sheetet átrakja a  legelsőre
            wb.Worksheets[1].Range["A1", (GetExcelColumnName(15)+ getRowNum(wb.Worksheets[1]))].UnMerge();

            for (int i = 2; i < wb.Worksheets.Count; i++) {
                int goalSheetRow = getRowNum(wb.Worksheets[1]); //első sheet utolsó használt sora
                int sourceSheetRow = getRowNum(wb.Worksheets[i]); //i. sheet utolsó használt sora
                string twentyEnd = GetExcelColumnName(15) + sourceSheetRow; //meddig akarjuk hogy unmergeltek legyenek a cellák?
                wb.Worksheets[i].Range["A1", twentyEnd].UnMerge(); //azért kell, mert a mergelt cellák gondot okozhatnak

                int sourceSheetColumn = getColNum(wb.Worksheets[i]); //i.sheet utolsó használt oszlopa
                
                string sourceSheetEnd = GetExcelColumnName(sourceSheetColumn+1) + sourceSheetRow;
                string goalSheetStart = "A" + (goalSheetRow+1);

                //Console.WriteLine(goalSheetStart+" "+sourceSheetRow+" "+sourceSheetColumn); //testing
                //Console.WriteLine(goalSheetStart); //testing
                wb.Worksheets[i].Range["A1", sourceSheetEnd].Copy(wb.Worksheets[1].Range[goalSheetStart]);
            }
            //wb.SaveCopyAs(@"D:\Szoftverfejlesztes-Targygraf\Book1.xlsx"); //testing, ha ki akarjuk menteni valamelyik sheetet
   
        }

        public string hunOnly(string text) //próbálkoztam a magyar és angol szöveg felezésével
        {
            string result;
            MatchCollection matches = Regex.Matches(text, @"( )");
            int count = matches.Count / 2+1;
            MatchCollection matches2 = Regex.Matches(text, @"(.*?\s){"+count+"}");
            result= matches2[0].Groups[0].Value;
            return result;
        }


        public void getSzak() {
            string szakNev = "";
            string szakTipus = "";
            string ervenyesseg = "";
         
            Regex nevFilter = new Regex(@"^(.*?)(BSc)|^(.*?)(MSc)");
            if (nevFilter.IsMatch(ws.Cells[1, 1].Value2.ToString())) //kiszedjük a szak nevét és tipusát
            {
                MatchCollection matches = Regex.Matches(ws.Cells[1, 1].Value2.ToString(), @"^(.*?)(BSc)|^(.*?)(MSc)");
                szakNev = matches[0].Groups[1].Value;
                szakTipus = matches[0].Groups[2].Value;
                Console.WriteLine(szakNev + " " + szakTipus); //testing
            }

            Regex ervenyessegFilter = new Regex(@"\w.*(\d{4}\/\d{2}.*){2}");
            if (ervenyessegFilter.IsMatch(ws.Cells[3, 1].Value2.ToString())) //kiszedjük a szak érvényességét mint szöveg
            {
                MatchCollection matches = Regex.Matches(ws.Cells[3, 1].Value2.ToString(), @"\w.*(\d{4}\/\d{2}.*){2}");
                ervenyesseg = matches[0].Groups[0].Value;
                Console.WriteLine(ervenyesseg); //testing
            }

        }

        public int getActualCol(int rowNum, int desiredCol) { //a mergelt cellák miatt lehet, hogy nem a várt cella tartalmazza az értéket
            int lastCol = getColNum(ws);
            //Console.WriteLine(sourceSheetColumn); //testing
            int actualCol = 1;
            int notEmptyCells = 0;
            while ((actualCol < lastCol) && (notEmptyCells < desiredCol))
            {
                if ((ws.Cells[(rowNum + 1), actualCol].Value2) != null) //ha a kategóriáknál nem üres az adott cella
                {
                    notEmptyCells++;
                }
                actualCol++;
            }
            actualCol--;
            //Console.WriteLine(actualCol); //testing

            return actualCol;
        }

        public int getUjFelev(int wsRowCount, int currentRow, string ajanlottFelev, string targyKategoria, ref string targyTipus) {
            int kreditCol = getActualCol(currentRow, 4);
            int elofeltetelCol = getActualCol(currentRow, 6);
            int targyKodCol = getActualCol(currentRow, 2);

            currentRow += 2;
            Regex osszesFilter = new Regex(@"^Összesen|^Összesítés|^Kreditpontok a modell ?tanterv féléveiben");
            Regex szabvalFilter = new Regex(@"(?i)^Tantárgy\sneve");

            while (currentRow < wsRowCount && ws.Cells[currentRow, 1].Value2 != null && !osszesFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //amig nincs vége a tárgyaknak
            {
                if (ws.Cells[currentRow, kreditCol].Value2 == null && ws.Cells[currentRow+1, 1].Value2 != null
                    && szabvalFilter.IsMatch(ws.Cells[currentRow + 1, 1].Value2.ToString())) //ha uj targykategoria van, akkor visszalépünk
                {
                    return --currentRow;
                }
                else if (ws.Cells[currentRow, kreditCol].Value2 == null) //ha uj targytipust talalunk
                {
                    targyTipus = ws.Cells[currentRow, 1].Value2.ToString();
                }
                else //ha uj tantargyat talalunk
                {
                    getUjTantargy(kreditCol, elofeltetelCol, targyKodCol, currentRow, ajanlottFelev, targyKategoria, targyTipus);
                }
                currentRow++;       
                
            }
            return --currentRow;
        }

        public void setTipusToDiff(ref int rowNum, ref string targyTipus) { //megváltoztatjuk a tárgytipust
            targyTipus = ws.Cells[rowNum, 1].Value2.ToString();
        }

       public void getUjTantargy(int kreditCol, int elofeltetelCol, int targyKodCol, int currentRow, string ajanlottFelev, string targyKategoria,
           string targyTipus) {
            string targyNev = "";
            string targyKod = "";
            int kredit = 0;

            targyNev = Regex.Replace(ws.Cells[currentRow, 1].Value2.ToString(), @"\r\n?|\n", " "); //tárgy neve
            if (ws.Cells[currentRow, targyKodCol].Value2 != null) //ha van kódja
            {
                targyKod = ws.Cells[currentRow, targyKodCol].Value2.ToString();
            }
            else targyKod = ""; //ha nincs kódja
            Console.WriteLine(targyNev + " " + targyKod); //testing
            

            MatchCollection match = Regex.Matches(ws.Cells[currentRow, kreditCol].Value2.ToString(), @"^[0-9]*");
            kredit = Convert.ToInt32(match[0].Groups[0].Value); //kredit
            Console.WriteLine(ajanlottFelev + " " + targyKategoria+ " "+targyTipus + " " + kredit); //testing

            List<string> elofeltetelkod = getElofeltetelek(currentRow, elofeltetelCol); //elofeltetelek

            //itt kell lementeni tárgyként
        }

        private List<string> getElofeltetelek(int currentRow, int elofeltetelCol) {
            Regex elofeltetelFilter = new Regex(@"\(*\(*[A-Z]+[0-9]+[a-zA-Z]+\)*\**|[0-9]+ ?kredit"); 
            List<string> elofeltetelkod = new List<string>();

            if (ws.Cells[currentRow, elofeltetelCol].Value2 != null &&
                elofeltetelFilter.IsMatch(ws.Cells[currentRow, elofeltetelCol].Value2.ToString()))
            {
                //Console.WriteLine(ws.Cells[j, elofeltetelCol].Value2.ToString()); //testing
                MatchCollection matches = Regex.Matches(ws.Cells[currentRow, elofeltetelCol].Value2.ToString(), @"\(*\(*[A-Z]+[0-9]+[a-zA-Z]+\)*\**|[0-9]+ ?kredit");
                 //\(*([A-Z]+[0-9]+[a-zA-Z]+)\)*\** ?
                foreach (Match elofeltetelmatch in matches)
                {
                    Console.WriteLine(elofeltetelmatch.Value); //testing
                    elofeltetelkod.Add(elofeltetelmatch.Value);
                }
            }
            return elofeltetelkod;
        }

        public void readInData()
        {
            getSzak();

            int felevCounter = 1;
            string ajanlottFelev = "";
            string targyKategoria = "Kötelező szakmai tárgyak";
            string targyTipus = "";
            int lastRow = getRowNum(ws);
            Regex osszesFilter = new Regex(@"^Összesítés|^Kreditpontok a modell ?tanterv féléveiben");
            Regex mscFelevFilter = new Regex(@"(.* félév[\s\S]*\d. félév[\s\S]*\d. félév[\s\S]*esetén\))");


            for (int currentRow = 1; currentRow < lastRow; currentRow++)
                {
                    targyTipus = "";
                    ajanlottFelev = Convert.ToString(felevCounter);
                    Regex felevFilter = new Regex(felevCounter + ". *félév");
                    Regex szabvalFilter = new Regex(@"(?i)^Tantárgy\sneve");

                if (ws.Cells[currentRow, 1].Value2 != null && osszesFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString()))
                    { //ha elértünk a végéhez
                        return;
                    }
                else if (ws.Cells[currentRow, 1].Value2 != null && mscFelevFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //ha uj msc felev van
                    {
                        MatchCollection matches = Regex.Matches(ws.Cells[currentRow, 1].Value2.ToString(), @"(.* félév[\s\S]*\d. félév[\s\S]*\d. félév[\s\S]*esetén\))");
                        ajanlottFelev = matches[0].Value;
                        currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, ref targyTipus);
                        felevCounter++;
                    }
                else if (ws.Cells[currentRow, 1].Value2 != null && felevFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //ha uj felev van
                    {
                        currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, ref targyTipus);
                        felevCounter++;
                    }
                else if ((felevCounter > 1 || felevCounter == 0) && ws.Cells[currentRow, 1].Value2 != null && ws.Cells[currentRow + 1, 1].Value2 != null &&
                        szabvalFilter.IsMatch(ws.Cells[currentRow + 1, 1].Value2.ToString())) //ha diff-es kategoria
                    {
                        felevCounter = 0;
                        ajanlottFelev = Convert.ToString(felevCounter);
                        setTipusToDiff(ref currentRow, ref targyKategoria);
                        currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, ref targyTipus);
                    }
                }

        }
    }
}
  