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
    public partial class Form4 : Form
    {
        List<Button> clicked = new List<Button>();
        public Form4()
        {
            InitializeComponent();
        }

        void Button_click(object sender, EventArgs e)
        {
            if (Program.admin == false)
            {
                Button button = (Button)sender;
                if (button.Text == "N")
                {
                    if (button.BackColor == System.Drawing.Color.Lime)
                    {
                        button.BackColor = Color.GreenYellow;
                        clicked.Add(button);


                    }
                    else
                    {
                        MessageBox.Show("This seat is already occupied");
                    }
                }
                else if (button.Text == "V")
                {
                    if (button.BackColor == System.Drawing.Color.Yellow)
                    {
                        button.BackColor = Color.Goldenrod;
                        clicked.Add(button);

                    }
                    else
                    {
                        MessageBox.Show("This seat is already occupied");
                    }
                }
            }
            else
            {
                MessageBox.Show("As an admin you are not allowed to purchase tickets");
            }
        }



        private void button209_Click(object sender, EventArgs e)
        {
            int normal_counter = 0;
            int vip_counter = 0;
            Button[] array = clicked.ToArray();
            int sum = 0;
            for (int i = 0; i < array.Length; i++)
            {
                if (clicked[i].BackColor == Color.GreenYellow)
                {
                    sum += 15000;
                    normal_counter++;
                }
                else if (clicked[i].BackColor == Color.Goldenrod)
                {
                    sum += 50000;
                    vip_counter++;
                }
                Program.all.Add(array[i]);
            }
            Tuple<int, string> temp1 = new Tuple<int, string>(Program.counter1, Program.logged_in);
            Program.counter1++;
            Tuple<int, int> temp = new Tuple<int, int>(normal_counter + vip_counter, sum);
            if (sum != 0)
            {
                Program.purchases1.Add(temp1, temp);
            }
            Program.match2_all_normal_counter += normal_counter;
            Program.match2_all_vip_counter += vip_counter;
            Program.match2_sum += sum;
            MessageBox.Show("The sum is:" + sum + "\t Normal tickets: " + normal_counter + "\t VIP tickets:" + vip_counter);
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i].BackColor == Color.GreenYellow)
                {
                    array[i].BackColor = Color.ForestGreen;
                }
                else if (array[i].BackColor == Color.Goldenrod)
                {
                    array[i].BackColor = Color.DarkKhaki;
                }
            }
            clicked.Clear();
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = null;
            }


        }

        private void button210_Click(object sender, EventArgs e)
        {
            Button[] array = clicked.ToArray();
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i].BackColor == Color.GreenYellow)
                {
                    array[i].BackColor = Color.Lime;
                }
                else if (array[i].BackColor == Color.Goldenrod)
                {
                    array[i].BackColor = Color.Yellow;
                }
            }
            clicked.Clear();
            for (int i = 0; i < array.Length; i++)
            {
                array[i] = null;
            }
        }

        private void button211_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            Owner.Show();
            this.Hide();
        }

        private void button212_Click(object sender, EventArgs e)
        {
            if (Program.admin == true)
            {
                foreach (KeyValuePair<Tuple<int, string>, Tuple<int, int>> kvp in Program.purchases1)
                {
                    textBox1.Text += kvp.Key.Item1 + ". Buyer: " + kvp.Key.Item2 + " Tickets: " + kvp.Value.Item1 + " Paid amount:" + kvp.Value.Item2+"\r\n" ;
                }
                textBox1.Text += "Value of normal tickets bought is:" + Program.match2_all_normal_counter * 15000 + "\tValue of vip tickets bought is:" + Program.match2_all_vip_counter * 50000 + "\r\n";
                textBox1.Text += "Value of all tickets bought is: " + Program.match2_sum;
            }
            else
            {
                MessageBox.Show("You need to be an admin to have access to this feature");
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }
    }
    }
