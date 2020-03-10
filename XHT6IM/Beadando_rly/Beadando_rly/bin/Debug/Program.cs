using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Beadando_rly
{
    static class Program
    {
        public static List<Button> all = new List<Button>();
        public static int counter = 1;
        public static int counter1 = 1;
        public static string logged_in;
        public static int match1_all_normal_counter=0;
        public static int match1_all_vip_counter=0;
        public static int match2_all_normal_counter=0;
        public static int match2_all_vip_counter=0;
        public static int match1_sum=0;
        public static int match2_sum=0;
        public static Dictionary<Tuple<int, string>, Tuple<int, int>> purchases = new Dictionary<Tuple<int, string>, Tuple<int, int>>();
        public static Dictionary<Tuple<int, string>, Tuple<int, int>> purchases1 = new Dictionary<Tuple<int, string>, Tuple<int, int>>();
        public static bool admin = false;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        
        }
    }
}
