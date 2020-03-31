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

        Regex szabvalFilter = new Regex(@"(?i)^Tantárgy\sneve");
        Regex mscFelevFilter = new Regex(@"(.* félév[\s\S]*\d. félév[\s\S]*\d. félév[\s\S]*esetén\))");
        Regex kommentFilter = new Regex(@"(\* ?([^*]+))+");
        
        Regex uresFilter = new Regex(@"-|^\s*$");

        Szak szak;
        List<Targy> targyak;
        List<Komment> kommentek = new List<Komment>();
        Dictionary<string, string[]> megjegyzesek;


        public excelFiles() {
            //wb = excel.Workbooks.Open(@"d:\szoftverfejlesztes\proginfo.xlsx");

            var allFiles = Directory.EnumerateFiles(@"D:\Szoftverfejlesztes-Targygraf\excelek");
            Regex fileFilter = new Regex(@"^(.)*?\\[^~$]+.xlsx$");
            foreach (string file in allFiles)
            {
                if (fileFilter.IsMatch(file)) //ha új excel fájlt találunk
                {
                    wb = excel.Workbooks.Open(file);

                    mergeSheets(); //az első sheetre rakunk mindent              
                    ws = wb.Worksheets[1]; //ws mindig az adott workbook első sheetje
                    targyak = new List<Targy>();
                    megjegyzesek = new Dictionary<string, string[]>();
                    readInData(); //beolvasunk
                    //addKomments();
                    //insertCurrentDatas();

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
            wb.SaveCopyAs(@"D:\Szoftverfejlesztes-Targygraf\Book1.xlsx"); //testing, ha ki akarjuk menteni valamelyik sheetet
   
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
         
            Regex nevFilter = new Regex(@"^(.*?) (BSc|MSc|felsőoktatási szakképzés|továbbképzési szak|BProf)");
            if (nevFilter.IsMatch(ws.Cells[1, 1].Value2.ToString())) //kiszedjük a szak nevét és tipusát
            {
                MatchCollection matches = Regex.Matches(ws.Cells[1, 1].Value2.ToString(), @"^(.*?) (BSc|MSc|felsőoktatási szakképzés|továbbképzési szak|BProf)");
                szakNev = matches[0].Groups[1].Value;
                szakTipus = matches[0].Groups[2].Value;
                //Console.WriteLine(szakNev + " " + szakTipus); //testing
            }

            Regex ervenyessegFilter = new Regex(@"\w.*(\d{4}\/\d{2}.*)"); // \w.*(\d{4}\/\d{2}.*){2}
            if (ervenyessegFilter.IsMatch(ws.Cells[3, 1].Value2.ToString())) //kiszedjük a szak érvényességét mint szöveg
            {
                MatchCollection matches = Regex.Matches(ws.Cells[3, 1].Value2.ToString(), @"\w.*(\d{4}\/\d{2}.*)");
                ervenyesseg = matches[0].Groups[0].Value;
                //Console.WriteLine(ervenyesseg); //testing
            }
            
            szak = new Szak(szakNev, szakTipus, ervenyesseg);
            szak.printSzak();
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

        public int getUjFelev(int wsRowCount, int currentRow, string ajanlottFelev, string targyKategoria, string targyTipus) {
            int kreditCol = getActualCol(currentRow, 4);
            int elofeltetelCol = getActualCol(currentRow, 6);
            int targyKodCol = getActualCol(currentRow, 2);
            currentRow += 2;
            Regex vegeFilter = new Regex(@"(?i)^Összesen|^Összesítés|^Kreditpontok a modell ?tanterv féléveiben");

            while (currentRow < wsRowCount && ws.Cells[currentRow, 1].Value2 != null && !vegeFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //amig nincs vége a tárgyaknak
            {
                if (ws.Cells[currentRow, kreditCol].Value2 == null && ws.Cells[currentRow + 1, 1].Value2 != null
                    && szabvalFilter.IsMatch(ws.Cells[currentRow + 1, 1].Value2.ToString())) //ha uj targykategoria van, akkor visszalépünk
                {
                    return --currentRow;
                }
                else if (ws.Cells[currentRow, kreditCol].Value2 == null && kommentFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString()))
                { //ha kommentet találtunk
                    MatchCollection matches = Regex.Matches(ws.Cells[currentRow, 1].Value2.ToString(), @"(\* ?([^*]+))+");
                    string[] array = new string[matches.Count];
                    for(int i=0; i<matches.Count; i++)
                    {
                        //kommentek.Add(new Komment(ajanlottFelev, match.Groups[2].Value));
                        //Console.WriteLine(match.Groups[2].Value); //testing
                        array[i] = matches[i].Groups[2].Value;
                        Console.WriteLine(matches[i].Groups[2].Value);

                    }
                    megjegyzesek.Add(ajanlottFelev, array);
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


       public void getUjTantargy(int kreditCol, int elofeltetelCol, int targyKodCol, int currentRow, string ajanlottFelev, string targyKategoria,
           string targyTipus) {
            string targyNev = "";
            string targyKod = "";
            int kredit = 0;
            string elofeltetelstring = "";

            targyNev = Regex.Replace(ws.Cells[currentRow, 1].Value2.ToString(), @"\r\n?|\n", " "); //tárgy neve
            if (ws.Cells[currentRow, targyKodCol].Value2 != null) //ha van kódja
            {
                targyKod = ws.Cells[currentRow, targyKodCol].Value2.ToString();
            }
            else targyKod = ""; //ha nincs kódja
            //Console.WriteLine(targyNev + " " + targyKod); //testing 

            MatchCollection match = Regex.Matches(ws.Cells[currentRow, kreditCol].Value2.ToString(), @"^[0-9]*");
            kredit = Convert.ToInt32(match[0].Groups[0].Value); //kredit
            //Console.WriteLine(ajanlottFelev + " " + targyKategoria+ " "+targyTipus + " " + kredit); //testing

            if (ws.Cells[currentRow, elofeltetelCol].Value2!=null && !uresFilter.IsMatch(ws.Cells[currentRow, elofeltetelCol].Value2.ToString()))
            { //elofeltetelek
                elofeltetelstring = ws.Cells[currentRow, elofeltetelCol].Value2.ToString();
            }
            else elofeltetelstring = "";

            targyak.Add(new Targy(targyNev, targyKod, ajanlottFelev, targyKategoria, targyTipus, kredit, elofeltetelstring));
            //itt kell lementeni tárgyként
        }

        //public void addKomments()
        //{
        //    foreach (Targy targy in targyak)
        //    {
        //        int starCount = targy.nev.Length - targy.nev.Replace("*", "").Length;
        //        if (starCount!=0) {
        //            targy.setMegjegyzes(megjegyzesek[targy.ajanlottfelev][starCount-1]);
        //            Console.WriteLine(megjegyzesek[targy.ajanlottfelev][starCount - 1]);
        //        }
        //        targy.nev = targy.nev.Replace("*", "");


        //    }
        //}

        public void addMegjegyzesek(SqliteDataAccess sqlite, string szakid) {
            foreach (var megjegyzes in megjegyzesek)
            {
                foreach (string text in megjegyzes.Value) {
                    try{
                        sqlite.ExecuteQuery("insert into Megjegyzés(szöveg) values ('" + text + "')");
                    }
                    catch (Exception e){
                        Console.WriteLine(e);
                    }           
                }
            }
            foreach (Targy targy in targyak)
            {
                int starCount = targy.nev.Length - targy.nev.Replace("*", "").Length;
                if (starCount != 0)
                {
                    try{
                        sqlite.ExecuteQuery("insert into Megjegyzése values (" + szakid + ", '" + targy.kod + "', '"
                        + megjegyzesek[targy.ajanlottfelev][starCount - 1] + "')");
                    }
                    catch (Exception e){
                        Console.WriteLine(e);
                    }
                    //targy.setMegjegyzes(megjegyzesek[targy.ajanlottfelev][starCount - 1]);                  
                    Console.WriteLine(megjegyzesek[targy.ajanlottfelev][starCount - 1]);
                }
                //targy.nev = targy.nev.Replace("*", "");
            }
        }


        public void insertCurrentDatas() {
            SqliteDataAccess sqlite = new SqliteDataAccess();
            szak.printSzak();
            sqlite.ExecuteQuery(szak.getInsert());

            string szakid = sqlite.QueryResult("select id from Szak where név='" + szak.szakNev +
                "' and [képzés típusa]='" + szak.szakTipus + "' and érvényesség='" + szak.ervenyesseg + "'");

            Console.WriteLine(szakid);
            foreach (Targy targy in targyak)
            {
                targy.printTargy();
                try{
                    sqlite.ExecuteQuery(targy.getInsertTantargy());
                }
                catch (Exception e){
                     Console.WriteLine(e);
                }
                try{
                    sqlite.ExecuteQuery(targy.getInsertTantargya(szakid));
                }
                catch (Exception e){
                    Console.WriteLine(e.Message);
                }
                if (targy.kreditfeltetel!=0) {
                    try{
                        sqlite.ExecuteQuery(targy.getInsertKreditfeltetel());
                    }
                    catch (Exception e){
                        Console.WriteLine(e.Message);
                    }
                    try{
                        sqlite.ExecuteQuery(targy.getInsertKreditFeltetele(szakid));
                    }
                    catch (Exception e){
                        Console.WriteLine(e.Message);
                    }
                }

            }
            addMegjegyzesek(sqlite, szakid);

            foreach (Targy targy in targyak)
            {
                foreach(Elofeltetel feltetel in targy.elofeltetelek)
                {
                    try{
                        sqlite.ExecuteQuery(feltetel.getInsertElofeltetele(szakid, targy.kod));
                    }
                    catch (Exception e){
                        Console.WriteLine(e.Message);
                    }
                }
                foreach (Elofeltetel feltetel in targy.elofeltetelek)
                {
                    if (!string.IsNullOrEmpty(feltetel.masodikkod)) {
                        try
                        {
                            sqlite.ExecuteQuery(feltetel.getInsertHelyettesitheto(szakid, targy.kod));
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }

                }
            }
        }

        public void readInData()
        {
            getSzak();

            int felevCounter = 1;
            string ajanlottFelev = "";
            string targyKategoria = "Kötelező szakmai tárgyak";
            string targyTipus = "";

            int lastRow = getRowNum(ws);
            Regex vegeFilter = new Regex(@"(?i)^Összesítés|^Kreditpontok a modell ?tanterv féléveiben");

            for (int currentRow = 1; currentRow < lastRow; currentRow++)
            {
                targyTipus = "";
                ajanlottFelev = Convert.ToString(felevCounter);
                Regex felevFilter = new Regex(felevCounter + ". *félév");              

                if (ws.Cells[currentRow, 1].Value2 != null && vegeFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString()))
                { //ha elértünk a végéhez
                    return;
                }
                else if (ws.Cells[currentRow, 1].Value2 != null && mscFelevFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //ha uj msc felev van
                {
                    MatchCollection matches = Regex.Matches(ws.Cells[currentRow, 1].Value2.ToString(), @"(.* félév[\s\S]*\d. félév[\s\S]*\d. félév[\s\S]*esetén\))");
                    ajanlottFelev = matches[0].Value;
                    currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, targyTipus);
                    felevCounter++;
                }
                else if (ws.Cells[currentRow, 1].Value2 != null && felevFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString())) //ha uj felev van
                {
                    currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, targyTipus);
                    felevCounter++;
                }
                else if ((felevCounter > 1 || felevCounter == 0) && ws.Cells[currentRow, 1].Value2 != null && ws.Cells[currentRow + 1, 1].Value2 != null &&
                        szabvalFilter.IsMatch(ws.Cells[currentRow + 1, 1].Value2.ToString())) //ha diff-es kategoria
                {
                    felevCounter = 0;
                    ajanlottFelev = Convert.ToString(felevCounter);
                    targyKategoria = ws.Cells[currentRow, 1].Value2.ToString();
                    currentRow = getUjFelev(lastRow, currentRow, ajanlottFelev, targyKategoria, targyTipus);
                }
                else if (ws.Cells[currentRow, 1].Value2 != null && kommentFilter.IsMatch(ws.Cells[currentRow, 1].Value2.ToString()))
                { //komment esetén
                    MatchCollection matches = Regex.Matches(ws.Cells[currentRow, 1].Value2.ToString(), @"(\* ?([^*]+))+");
                    string[] array = new string[matches.Count];
                    for (int i = 0; i < matches.Count; i++)
                    {
                        //kommentek.Add(new Komment(ajanlottFelev, match.Groups[2].Value));
                        //Console.WriteLine(match.Groups[2].Value); //testing
                        array[i] = matches[i].Groups[2].Value;
                        Console.WriteLine(matches[i].Groups[2].Value);

                    }
                    megjegyzesek.Add(ajanlottFelev, array);
                }
            }

        }
    }
}
  