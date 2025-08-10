using System.Windows.Forms;

namespace TitlePageGenerator
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private Button btnTitle;
        private Button btnCover;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnTitle = new System.Windows.Forms.Button();
            this.btnCover = new System.Windows.Forms.Button();
            this.SuspendLayout();

            this.btnTitle.Location = new System.Drawing.Point(50, 30);
            this.btnTitle.Name = "btnTitle";
            this.btnTitle.Size = new System.Drawing.Size(200, 40);
            this.btnTitle.Text = "Титульный лист";
            this.btnTitle.UseVisualStyleBackColor = true;
            this.btnTitle.Click += new System.EventHandler(this.btnTitle_Click);

            this.btnCover.Location = new System.Drawing.Point(50, 80);
            this.btnCover.Name = "btnCover";
            this.btnCover.Size = new System.Drawing.Size(200, 40);
            this.btnCover.Text = "Заявочный лист";
            this.btnCover.UseVisualStyleBackColor = true;
            this.btnCover.Click += new System.EventHandler(this.btnCover_Click);
 
            this.ClientSize = new System.Drawing.Size(300, 160);
            this.Controls.Add(this.btnCover);
            this.Controls.Add(this.btnTitle);
            this.Name = "MainForm";
            this.Text = "Генератор документов";
            this.ResumeLayout(false);
        }
    }
}