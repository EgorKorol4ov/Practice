using System;
using System.Windows.Forms;

namespace TitlePageGenerator
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnTitle_Click(object sender, EventArgs e)
        {
            frmTitle f = new frmTitle();
            f.ShowDialog();
        }

        private void btnCover_Click(object sender, EventArgs e)
        {
            frmCover f = new frmCover();
            f.ShowDialog();
        }
    }
}