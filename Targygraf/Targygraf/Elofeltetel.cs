using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Targygraf
{
    class Elofeltetel
    {
        public string elsokod;
        public int egyszerrefelveheto;
        public string masodikkod;

        public Elofeltetel(string kod, int egyszerrefelveheto, string helyettesitokod)
        {
            this.elsokod = kod;
            this.egyszerrefelveheto = egyszerrefelveheto;
            this.masodikkod = helyettesitokod;
        }

        public void printElofeltetel() {
            Console.WriteLine("Elofeltetele: "+ elsokod+ ", " +egyszerrefelveheto+", " +masodikkod);
        }

        public string getInsertElofeltetele(string szakid, string targykod) { 
            return "insert into Előfeltétele(előfeltételkód, ráépülőkód, szakid, egyszerrefelveheto) values " +
                              "('" + elsokod + "', '" + targykod + "', "+szakid+", "+egyszerrefelveheto+")";
        }

        public string getInsertHelyettesitheto(string szakid, string targykod)
        {
            return "insert into Helyettesíthető(elsőkód, elsőráépülő, elsőszak, másodikkód, másodikráépülő, másodikszak) values " +
                              "('" + elsokod + "', '" + targykod + "', " + szakid + ", '" 
                              + masodikkod + "', '"+targykod+"', "+szakid+")";
        }
    }
}
