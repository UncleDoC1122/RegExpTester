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
            List<string> cols = new List<string>();

            List<List<string>> data = new List<List<string>>();

            cols.Add("A");
            data.Add(new List<string>());
            data[0].Add("1");
            data[0].Add("Test");
            WorkWithExcel.putColumns("D:\\Test.xlsx", 1, cols, data);
        }
    }
}
