namespace TitlePageGenerator
{
    partial class frmCover
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
            // 
            // createDocButton
            // 
            this.createDocButton.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.createDocButton.Location = new System.Drawing.Point(213, 31);
            this.createDocButton.Margin = new System.Windows.Forms.Padding(4);
            this.createDocButton.Name = "createDocButton";
            this.createDocButton.Size = new System.Drawing.Size(738, 60);
            this.createDocButton.TabIndex = 0;
            this.createDocButton.Text = "Создать документ";
            this.createDocButton.UseVisualStyleBackColor = true;
            this.createDocButton.Click += new System.EventHandler(this.createDocButton_Click);
            // 
            // frmCover
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1157, 120);
            this.Controls.Add(this.createDocButton);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmCover";
            this.ResumeLayout(false);

        }
    }
}