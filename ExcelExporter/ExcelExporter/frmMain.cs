using System;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

            InitializeDataGridView();
        }

        private static readonly string[] ColumnHeaders =
        {
            "№",
            "Наименование дисциплин",
            "Курс",
            "Семестр 1",
            "Семестр 2",
            "Шифр группы",
            "Контингент",
            "Кол-во потоков",
            "Кол-во групп",
            "Кол-во подгрупп",
            "Лекц. на 1 гр.",
            "Лекц. всего",
            "Практ. на 1 гр.",
            "Практ. всего",
            "Лаб. на 1 гр.",
            "Лаб. всего",
            "ГАК, ДП",
            "Летняя практика",
            "Рук. аспир.",
            "Доп. обр. структуры",
            "Всего акад. час.",
            "Всего в кредитах"
        };

        private void InitializeDataGridView()
        {
            dgvData.ColumnCount = ColumnHeaders.Length;

            for (int i = 0; i < ColumnHeaders.Length; i++)
            {
                dgvData.Columns[i].Name = ColumnHeaders[i];
            }

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

        private static readonly double[] ColumnWidths =
        {
            3.3,
            18.4,
            2.9,
            6,
            6,
            8.5,
            4.9,
            4.9,
            4.9,
            4.9,
            4.4,
            4.4,
            4.4,
            4.4,
            4.4,
            4.4,
            6.6,
            5,
            5,
            8.5,
            5,
            5.1
        };

        private void SetColumnWidths(Excel.Worksheet sheet)
        {
            int startIndex = 4;
            for (int i = 0; i < ColumnWidths.Length; i++)
            {
                sheet.Columns[startIndex + i].ColumnWidth = ColumnWidths[i];
            }
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
