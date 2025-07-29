using System;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            dgvData.ColumnCount = 22;

            dgvData.Columns[0].Name = "№";
            dgvData.Columns[1].Name = "Наименование дисциплин";
            dgvData.Columns[2].Name = "Курс";
            dgvData.Columns[3].Name = "Семестр 1";
            dgvData.Columns[4].Name = "Семестр 2";
            dgvData.Columns[5].Name = "Шифр группы";
            dgvData.Columns[6].Name = "Контингент";
            dgvData.Columns[7].Name = "Кол-во потоков";
            dgvData.Columns[8].Name = "Кол-во групп";
            dgvData.Columns[9].Name = "Кол-во подгрупп";
            dgvData.Columns[10].Name = "Лекц. на 1 гр.";
            dgvData.Columns[11].Name = "Лекц. всего";
            dgvData.Columns[12].Name = "Практ. на 1 гр.";
            dgvData.Columns[13].Name = "Практ. всего";
            dgvData.Columns[14].Name = "Лаб. на 1 гр.";
            dgvData.Columns[15].Name = "Лаб. всего";
            dgvData.Columns[16].Name = "ГАК, ДП";
            dgvData.Columns[17].Name = "Летняя практика";
            dgvData.Columns[18].Name = "Рук. аспир.";
            dgvData.Columns[19].Name = "Доп. обр. структуры";
            dgvData.Columns[20].Name = "Всего акад. час.";
            dgvData.Columns[21].Name = "Всего в кредитах";

            dgvData.AllowUserToAddRows = true;
            dgvData.AllowUserToDeleteRows = true;
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            dgvData.Rows.Add();
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            object missing = Missing.Value;
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Add(missing);
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            try
            {
                sheet.Name = "Академическая ведомость";

                CreateComplexHeader(sheet);

                SetColumnWidths(sheet);

                FillDataFromGrid(sheet);

                ApplyBordersAndFormatting(sheet);

                string filePath = Application.StartupPath + @"\AcademicTable.xlsx";
                workbook.SaveAs(filePath);
                workbook.Close(false, missing, missing);
                excelApp.Quit();

                MessageBox.Show("Файл успешно сохранён:\n" + filePath, "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
                workbook.Close(false, missing, missing);
                excelApp.Quit();
            }
        }

        private void FillDataFromGrid(Excel.Worksheet sheet)
        {
            int startRow = 8;

            for (int i = 0; i < dgvData.Rows.Count; i++)
            {
                if (!dgvData.Rows[i].IsNewRow)
                {
                    for (int j = 0; j < dgvData.Columns.Count; j++)
                    {
                        var cellValue = dgvData.Rows[i].Cells[j].Value;
                        sheet.Cells[startRow + i, j + 4] = cellValue?.ToString() ?? "";
                    }
                }
            }
        }

        private void CreateComplexHeader(Excel.Worksheet sheet)
        {
            sheet.Range["D5", "D7"].Merge();
            sheet.Range["D5"].Value = "№";

            sheet.Range["E5", "E7"].Merge();
            sheet.Range["E5"].Value = "Наименование дисциплин";

            sheet.Range["F5", "F7"].Merge();
            sheet.Range["F5"].Value = "Курс";
            sheet.Range["F5", "F7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["G5", "H5"].Merge();
            sheet.Range["G5"].Value = "Семестр";

            sheet.Range["G6", "G7"].Merge();
            sheet.Range["G6"].Value = "1";

            sheet.Range["H6", "H7"].Merge();
            sheet.Range["H6"].Value = "2";

            sheet.Range["I5", "I7"].Merge();
            sheet.Range["I5"].Value = "Шифр группы";
            sheet.Range["I5", "I7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["J5", "J7"].Merge();
            sheet.Range["J5"].Value = "Контингент";
            sheet.Range["J5", "J7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["K5", "K7"].Merge();
            sheet.Range["K5"].Value = "Кол-во потоков";
            sheet.Range["K5", "K7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["L5", "L7"].Merge();
            sheet.Range["L5"].Value = "Кол-во групп";
            sheet.Range["L5", "L7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["M5", "M7"].Merge();
            sheet.Range["M5"].Value = "Кол-во подгрупп";
            sheet.Range["M5", "M7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["N5", "S5"].Merge();
            sheet.Range["N5"].Value = "Вид занятий";

            sheet.Range["N6", "O6"].Merge();
            sheet.Range["N6"].Value = "лекц.";

            sheet.Range["P6", "Q6"].Merge();
            sheet.Range["P6"].Value = "практ. зан.";

            sheet.Range["R6", "S6"].Merge();
            sheet.Range["R6"].Value = "лаб. зан.";

            sheet.Range["N7"].Value = "на 1 гр.";
            sheet.Range["N7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["O7"].Value = "всего";
            sheet.Range["O7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["P7"].Value = "на 1 гр.";
            sheet.Range["P7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["Q7"].Value = "всего";
            sheet.Range["Q7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["R7"].Value = "на 1 гр.";
            sheet.Range["R7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["S7"].Value = "всего";
            sheet.Range["S7"].Orientation = Excel.XlOrientation.xlUpward;


            sheet.Range["T5", "T7"].Merge();
            sheet.Range["T5"].Value = "ГАК, ДП";
            sheet.Range["T5", "T7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["U5", "U7"].Merge();
            sheet.Range["U5"].Value = "Летняя практика";
            sheet.Range["U5", "U7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["V5", "V7"].Merge();
            sheet.Range["V5"].Value = "Рук. аспир.";
            sheet.Range["V5", "V7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["W5", "W7"].Merge();
            sheet.Range["W5"].Value = "Дополнительные образовательные структуры";
            sheet.Range["W5", "W7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["X5", "X7"].Merge();
            sheet.Range["X5"].Value = "Всего акад. час.";
            sheet.Range["X5", "X7"].Orientation = Excel.XlOrientation.xlUpward;

            sheet.Range["Y5", "Y7"].Merge();
            sheet.Range["Y5"].Value = "Всего в кредит часах";
            sheet.Range["Y5", "Y7"].Orientation = Excel.XlOrientation.xlUpward;
        }

        private void SetColumnWidths(Excel.Worksheet sheet)
        {
            sheet.Columns["D"].ColumnWidth = 3.3;
            sheet.Columns["E"].ColumnWidth = 18.4;
            sheet.Columns["F"].ColumnWidth = 2.9;
            sheet.Columns["G"].ColumnWidth = 6;
            sheet.Columns["H"].ColumnWidth = 6;
            sheet.Columns["I"].ColumnWidth = 8.5;
            sheet.Columns["J"].ColumnWidth = 4.9;
            sheet.Columns["K"].ColumnWidth = 4.9;
            sheet.Columns["L"].ColumnWidth = 4.9;
            sheet.Columns["M"].ColumnWidth = 4.9;
            sheet.Columns["N"].ColumnWidth = 4.4;
            sheet.Columns["O"].ColumnWidth = 4.4;
            sheet.Columns["P"].ColumnWidth = 4.4;
            sheet.Columns["Q"].ColumnWidth = 4.4;
            sheet.Columns["R"].ColumnWidth = 4.4;
            sheet.Columns["S"].ColumnWidth = 4.4;
            sheet.Columns["T"].ColumnWidth = 6.6;
            sheet.Columns["U"].ColumnWidth = 5;
            sheet.Columns["V"].ColumnWidth = 5;
            sheet.Columns["W"].ColumnWidth = 8.5;
            sheet.Columns["X"].ColumnWidth = 5;
            sheet.Columns["Y"].ColumnWidth = 5.1;
        }

        private void ApplyBordersAndFormatting(Excel.Worksheet sheet)
        {
            Excel.Range headerRange = sheet.get_Range("D5", "Y7");
            headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            headerRange.Borders.Weight = 3;

            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.WrapText = true;

            headerRange.Font.Name = "Times New Roman";
            headerRange.Font.Size = 12;
            headerRange.Font.Bold = false;

            sheet.Rows[5].RowHeight = 19.5;
            sheet.Rows[6].RowHeight = 16;
            sheet.Rows[7].RowHeight = 63;

            int dataRowCount = dgvData.Rows.Count;
            if (dataRowCount > 0)
            {
                int lastDataRow = 8 + dataRowCount - 1;

                int actualDataRows = 0;
                for (int i = 0; i < dgvData.Rows.Count; i++)
                {
                    if (!dgvData.Rows[i].IsNewRow)
                    {
                        actualDataRows++;
                    }
                }

                if (actualDataRows > 0)
                {
                    lastDataRow = 8 + actualDataRows - 1;
                    Excel.Range dataRange = sheet.get_Range("D8", "Y" + lastDataRow);

                    dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    dataRange.Borders.Weight = 3;

                    dataRange.Font.Name = "Times New Roman";
                    dataRange.Font.Size = 12;
                    dataRange.Font.Bold = false;

                    dataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    dataRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
            }
        }
    }
}
