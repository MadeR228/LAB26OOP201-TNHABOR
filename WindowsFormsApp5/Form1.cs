using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click (object sender, EventArgs e)
        {
            var helper = new Class1("wordcheck.docx");

            double summ1 = double.Parse(textBox15.Text);
            double summ2 = double.Parse(textBox16.Text);
            double summ3 = double.Parse(textBox17.Text);

            double generalSumm = summ1 + summ2 + summ3;

            var items = new Dictionary<string, string>
            {
                {"<COMPANYNAME>",textBox1.Text },
                {"<NMBFACTURA>",textBox2.Text}, 
                {"<DATE>", dateTimePicker1.Value.ToString("dd.MM.yyyy")},
                {"<FORWHAT>", textBox4.Text},
                {"<STREETCOMPANY>",textBox5.Text},
                {"<CITYCOMPANY>",textBox6.Text},
                {"<NMBOFREQUEST>",textBox7.Text},
                {"<NAME>",textBox8.Text},
                {"<COMPANY>",textBox9.Text},
                {"<STREET>",textBox10.Text},
                {"<CITY>",textBox11.Text},
                {"<MOBILE>",textBox12.Text},
                {"<DESCRIPTION1>",textBox3.Text},
                {"<DESCRIPTION2>", textBox13.Text},
                {"<DESCRIPTION3>",textBox14.Text},
                {"<SUMM1>",textBox15.Text},
                {"<SUMM2>",textBox16.Text},
                {"<SUMM3>",textBox17.Text},
                {"<GENERALSUMM>",generalSumm.ToString()}
            };
            helper.Process(items);
        }
    }
}
