namespace ExcelExporter
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnExportToExcel;
        private System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.Button btnAddRow;

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
            this.btnExportToExcel = new System.Windows.Forms.Button();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.btnAddRow = new System.Windows.Forms.Button();

            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.SuspendLayout();

            this.btnExportToExcel.Location = new System.Drawing.Point(12, 12);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(150, 40);
            this.btnExportToExcel.TabIndex = 0;
            this.btnExportToExcel.Text = "Выгрузка в Excel";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);

            this.btnAddRow.Location = new System.Drawing.Point(170, 12);
            this.btnAddRow.Name = "btnAddRow";
            this.btnAddRow.Size = new System.Drawing.Size(150, 40);
            this.btnAddRow.TabIndex = 2;
            this.btnAddRow.Text = "Добавить строку";
            this.btnAddRow.UseVisualStyleBackColor = true;
            this.btnAddRow.Click += new System.EventHandler(this.btnAddRow_Click);

            this.dgvData.AllowUserToAddRows = true;
            this.dgvData.AllowUserToDeleteRows = true;
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Location = new System.Drawing.Point(12, 60);
            this.dgvData.Name = "dgvData";
            this.dgvData.Size = new System.Drawing.Size(950, 400);
            this.dgvData.TabIndex = 1;

            this.ClientSize = new System.Drawing.Size(980, 480);
            this.Controls.Add(this.btnExportToExcel);
            this.Controls.Add(this.btnAddRow);
            this.Controls.Add(this.dgvData);
            this.Name = "Form1";
            this.Text = "Excel Выгрузка";

            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.ResumeLayout(false);
        }
    }
}
