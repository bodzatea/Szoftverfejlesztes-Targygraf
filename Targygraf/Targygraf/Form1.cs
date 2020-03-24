using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Targygraf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {  
        
        }

        public void OpenFile() {
            excelFiles excel = new excelFiles();
            label2.Text = "End";
            //excel.readTables();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            OpenFile();
          

            //SqliteDataAccess sql = new SqliteDataAccess();
            //sql.ExecuteQuery("insert into Targy(id, nev, kod, kredit) values(, 'Matek', 'VEMIMAB', 6)");
        }
    }
}
