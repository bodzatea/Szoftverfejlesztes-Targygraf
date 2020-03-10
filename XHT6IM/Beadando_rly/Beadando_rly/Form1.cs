using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace Beadando_rly
{
    public partial class Form1 : Form
    {
        Form2 f2;
        Match_Chooser f3;

        DB adatb = new DB();
        SQLiteCommand cmd;
        SQLiteDataAdapter sda;
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
                sda = new SQLiteDataAdapter("Select count(*) from users where username='" + textBox1.Text + "'and password = '" + textBox2.Text + "'", adatb.GetConnection());
                dt = new DataTable();
                sda.Fill(dt);
            bool found = false;
            if (dt.Rows[0][0].ToString() == "1")
            {
                found = true;
                if (f3 == null)
                {
                    f3 = new Match_Chooser();
                    f3.FormClosed += f3_FormClosed;
                }
                Program.logged_in = textBox1.Text;
                textBox1.Text = "";
                textBox2.Text = "";
                f3.Show(this);
                this.Hide();
            }
            else if (found == false)
            {
                sda = new SQLiteDataAdapter("Select count(*) from admins where username = '" + textBox1.Text + "' and password = '" + textBox2.Text + "'", adatb.GetConnection());
                dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    if (f3 == null)
                    {
                        f3 = new Match_Chooser();
                        f3.FormClosed += f3_FormClosed;
                    }
                    Program.logged_in = textBox1.Text;
                    textBox1.Text = "";
                    textBox2.Text = "";
                    Program.admin = true;
                    f3.Show(this);
                    this.Hide();

                }
            }
        }


        void f3_FormClosed(object sender, FormClosedEventArgs e)
        {
            f3 = null;  //If form is closed make sure reference is set to null
            Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This program was created by:Lőrincz Beáta(XHT6IM)");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
    }
