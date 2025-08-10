using System;
using System.Windows.Forms;

namespace TitlePageGenerator
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnTitle_Click(object sender, EventArgs e)
        {
            var f = new Form1();
            f.ShowDialog();
        }

        private void btnCover_Click(object sender, EventArgs e)
        {
            var f = new Form2();
            f.ShowDialog();
        }
    }
}