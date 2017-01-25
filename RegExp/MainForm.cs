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
    public partial class MainForm : Form
    {
        private string filepath;
        public MainForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                filepath = ofd.FileName;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> tmp = new List<string>();
            tmp.Add(textBox1.Text);
            List<List<string>> outp = RegExp.parseAddress(tmp);
        }
    }
}
