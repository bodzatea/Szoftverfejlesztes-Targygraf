using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Targygraf
{
    class Szak
    {
        public string szakNev;
        public string szakTipus;
        public string ervenyesseg;

        public Szak(string szakNev, string szakTipus, string ervenyesseg) {
            this.szakNev = szakNev;
            this.szakTipus = szakTipus;
            this.ervenyesseg = ervenyesseg;
        }

        public void printSzak() {
            Console.WriteLine(szakNev+ " "+ szakTipus);
            Console.WriteLine(ervenyesseg);
        }

        public void insertSzak() {
            SqliteDataAccess sqlite = new SqliteDataAccess();
            sqlite.ExecuteQuery(getInsert());
        }

        public string getInsert() {
            //return "if not exists (select 1 from Szak " +
            //    "where név='"+szakNev+ "' and [képzés típusa]='"+szakTipus+"' and érvényesség='"+ervenyesseg+"') " +
            //    "begin" +
            //    " insert into Szak(név, [képzés típusa], érvényesség) values " + 
            //                   "('" + szakNev + "', '" + szakTipus + "', '" + ervenyesseg + "') end;";
            return " insert into Szak(név, [képzés típusa], érvényesség) select " +
                   "'" + szakNev + "', '" + szakTipus + "', '" + ervenyesseg + "' " +
                   "where not exists (select 1 from Szak " +
                    "where név='" + szakNev + "' and [képzés típusa]='" + szakTipus + "' and érvényesség='" + ervenyesseg + "'); ";
        }
    }
}
