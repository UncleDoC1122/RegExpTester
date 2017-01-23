using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RegExp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> output = WorkWithExcel.parseExcel("D:\\Users\\DSIYANCHEV\\Desktop\\Нормализация таблиц\\Авто\\ADDRSTREET - возможна автоматическая обработка.xlsx", 1, "AG");

            for (int i = 0; i < output.Count; i ++)
            {
                textBox1.Text += output[i];
            }
        }
    }
}
