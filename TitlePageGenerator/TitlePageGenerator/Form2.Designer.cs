namespace TitlePageGenerator
{
    partial class Form2
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button createDocButton;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.createDocButton = new System.Windows.Forms.Button();
            this.SuspendLayout();

            this.createDocButton.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.createDocButton.Location = new System.Drawing.Point(30, 30);
            this.createDocButton.Name = "createDocButton";
            this.createDocButton.Size = new System.Drawing.Size(200, 40);
            this.createDocButton.TabIndex = 0;
            this.createDocButton.Text = "Создать документ";
            this.createDocButton.UseVisualStyleBackColor = true;
            this.createDocButton.Click += new System.EventHandler(this.createDocButton_Click);

            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(280, 100);
            this.Controls.Add(this.createDocButton);
            this.Name = "Form1";
            this.Text = "Генератор заявочного листа";
            this.ResumeLayout(false);
        }
    }
}