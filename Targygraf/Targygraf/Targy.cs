using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Targygraf
{ 
    class Targy
    {
        public string nev;
        public string kod;

        public string ajanlottfelev;
        string kategoria;
        string tipus;
        int kredit;
        string megjegyzes;

        string elofeltetelstring;
        public int kreditfeltetel;

        public List<Elofeltetel> elofeltetelek;

        public Targy(string nev, string kod, string ajanlottfelev, string kategoria, string tipus, 
            int kredit, string elofeltetelstring) 
        {
            this.nev = nev;
            this.kod = kod.ToUpper();
           
            this.kategoria = kategoria;
            this.tipus = tipus;
            this.ajanlottfelev = ajanlottfelev;
            this.kredit = kredit;
            this.megjegyzes = "";
            this.kreditfeltetel = 0;

            this.elofeltetelstring = elofeltetelstring;

            printTargy();
            kodFormat();

        }

        public void setMegjegyzes(string text)
        {
            this.megjegyzes = text;
        }

        public void printTargy() {
            Console.WriteLine(nev+" "+kod);
            Console.WriteLine(ajanlottfelev+", "+kategoria+" - "+tipus+ ", "+kredit);
            Console.WriteLine(elofeltetelstring);
        }

        public string getInsertTantargy() {
            string tempNev = nev.Replace("*", "");
            if (string.IsNullOrEmpty(kod))
            {
                return "insert into Kategória(név) values " +
                             "('" + tempNev + "')";
            }
            else {
                return "insert into Tantárgy(név, kód) values " +
                             "('" + tempNev + "', '" + kod + "')";
            }              
        }

        public string getInsertTantargya(string szakid) {
            if (string.IsNullOrEmpty(kod))
            {
                return "insert into Kategóriája(név, szakid, kredit, ajánlottfélév) values" +
               "('" + nev + "', " + szakid + ", " + kredit + ", '" + ajanlottfelev+ "')";
            }
            else {
                if (string.IsNullOrEmpty(tipus))
                {
                    tipus = "NULL";
                }
                else tipus="'"+tipus+"'";

                return "insert into Tantárgya(tantárgykód, szakid, kategória, típus, ajánlottfélév, kredit) values" +
               "('" + kod + "', " + szakid + ", '" + kategoria + "', " + tipus + ", '" + ajanlottfelev + "', " + kredit + ")";
            }
        }

        public string getInsertKreditfeltetel() {
                return "insert into Kreditfeltétel(minimumkredit) values" +
               "("+kreditfeltetel+")";
            
        }

        public string getInsertKreditFeltetele(string szakid) {
            return "insert into Kreditfeltétele(minimumkredit, tantárgykód, szakid) values" +
            "(" + kreditfeltetel + ", '"+kod+"', "+szakid+")";
        }




        private void kodFormat()
        {
            Regex kodFilter = new Regex(@"(\()?([A-Z]+[0-9]+[a-zA-Z]+)(\))?\*?\s*(vagy)?");
            Regex kreditFilter = new Regex(@"([0-9]+) *kredit");
            elofeltetelek = new List<Elofeltetel>();

                if (kreditFilter.IsMatch(elofeltetelstring)) {
                    //megjegyzésbe rakjuk a feltételt, esetleg kiszámoljuk?
                    MatchCollection matches = Regex.Matches(elofeltetelstring, @"([0-9]+) *kredit");
                    kreditfeltetel = Convert.ToInt32(matches[0].Groups[1].Value);
                    //Console.WriteLine(kreditfeltetel);
                }

                if (kodFilter.IsMatch(elofeltetelstring)) {
                    //berakjuk előfeltételbe, 
                    MatchCollection matches = Regex.Matches(elofeltetelstring, @"(\()?([A-Z]+[0-9]+[a-zA-Z]+)(\))?\*?\s*(vagy)?");
                    int egyszerrefelveheto = 0;
                    string helyettesitokod = "";

                    foreach (Match match in matches)
                    {
                        egyszerrefelveheto = 0;
                        if (match.Groups[1].Success && match.Groups[3].Success) {
                            //az előfeltétellel együtt felvehető a tantárgy
                            egyszerrefelveheto = 1;
                        }
                        elofeltetelek.Add(new Elofeltetel(match.Groups[2].Value, egyszerrefelveheto, helyettesitokod));

                        if (match.Groups[4].Success)
                        {
                            helyettesitokod = match.Groups[2].Value;
                            //a következő kóddal helyettesithető a feltétel
                        }
                        else helyettesitokod = "";

                        //Console.WriteLine(match.Value); //testing
                    }
                }
            }
    }


}
