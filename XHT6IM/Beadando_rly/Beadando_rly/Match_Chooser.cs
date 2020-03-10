using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Beadando_rly
{
    public partial class Match_Chooser : Form
    {
        Form2 f2;
        Form1 f1;
        Form4 f4;
        public Match_Chooser()
        {
            InitializeComponent();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex == 0)
            {
                if (f2 == null)
                {
                    f2 = new Form2();
                    f2.FormClosed += f2_FormClosed;
                }
                f2.Show(this);
                this.Hide();
            }
            else if (listBox1.SelectedIndex == 1)
            {
                if (f4 == null)
                {
                    f4 = new Form4();
                    f4.FormClosed += f4_FormClosed;
                }
                f4.Show(this);
                this.Hide();
            }
            else
            {
                MessageBox.Show("Nothing is selected");
            }
        }

        void f2_FormClosed(object sender, FormClosedEventArgs e)
        {
            f2 = null;  //If form is closed make sure reference is set to null
            Show();
        }
        void f1_FormClosed(object sender, FormClosedEventArgs e)
        {
            f1 = null;  //If form is closed make sure reference is set to null
            Show();
        }
        void f4_FormClosed(object sender, FormClosedEventArgs e)
        {
            f4 = null;  //If form is closed make sure reference is set to null
            Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Program.logged_in = "";
            if (Program.admin == true)
            {
                Program.admin = false;   
            }
            Owner.Show();
            this.Hide();
        }

        private void Match_Chooser_Load(object sender, EventArgs e)
        {

        }
    }
}
