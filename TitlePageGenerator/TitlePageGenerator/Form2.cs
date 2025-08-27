using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace TitlePageGenerator
{
    public partial class frmCover : Form
    {
        public frmCover()
        {
            InitializeComponent();
        }

        private void createDocButton_Click(object sender, EventArgs e)
        {
            GenerateDocument();
        }

        private void GenerateDocument()
        {
            try
            {
                Word.Application wordApp = new Word.Application();
                Word.Document doc = wordApp.Documents.Add();

                doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1f);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1f);
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2f);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(2f);

                wordApp.Visible = false;
                doc.Activate();

                AddCustomParagraph(doc, "Заявочный лист на участие в соревнованиях", "Times New Roman",
                16, 1, 1, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphCenter,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Команда ____________________________________", "Times New Roman",
                14, 1, 0, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "(название команды)", "Times New Roman",
                10, 0, 1, 0f, 5f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineSingle);

                AddCustomParagraph(doc, "Контактное лицо _____________________", "Times New Roman",
                14, 1, 0, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Образовательное учреждение, округ, адрес, телефон _________________________________________", "Times New Roman",
                14, 1, 0, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "e-mail _________________________________________________", "Times New Roman",
                14, 1, 0, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Многоборье", "Times New Roman",
                14, 1, 1, 0f, 0f, 0f, 0f, Word.WdParagraphAlignment.wdAlignParagraphCenter,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineSingle);

                Word.Table table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, 7, 7);
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                table.AllowAutoFit = false;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = wordApp.CentimetersToPoints(18);

                table.Borders.Enable = 1;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
                table.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;

                string[] headers =
                {
                    "№\nп/п",
                    "Фамилия, имя гимнастки",
                    "Категория",
                    "Возраст",
                    "Разряд",
                    "Тренер гимнастки",
                    "Отметка врача о допуске"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Range headerRange = table.Cell(1, i + 1).Range;
                    headerRange.Text = headers[i];

                    headerRange.Font.Name = "Times New Roman";
                    headerRange.Font.Size = 12;
                    headerRange.Font.Bold = 1;
                    headerRange.Font.Italic = 0;
                    headerRange.Font.Underline = Word.WdUnderline.wdUnderlineNone;

                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.ParagraphFormat.SpaceAfter = 0f;
                    headerRange.ParagraphFormat.SpaceBefore = 0f;
                    headerRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                    table.Cell(1, i + 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }

                float rowHeight = wordApp.CentimetersToPoints(1f);
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    table.Rows[row].Height = rowHeight;
                    table.Rows[row].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                }

                table.Columns[1].Width = wordApp.CentimetersToPoints(1f);
                table.Columns[2].Width = wordApp.CentimetersToPoints(6f);
                table.Columns[3].Width = wordApp.CentimetersToPoints(2.5f);
                table.Columns[4].Width = wordApp.CentimetersToPoints(2.5f);
                table.Columns[5].Width = wordApp.CentimetersToPoints(2.5f);
                table.Columns[6].Width = wordApp.CentimetersToPoints(4.3f);
                table.Columns[7].Width = wordApp.CentimetersToPoints(5f);

                for (int row = 2; row <= 7; row++)
                {
                    Word.Range numberCell = table.Cell(row, 1).Range;
                    numberCell.Text = (row - 1).ToString() + ".";

                    numberCell.Font.Name = "Times New Roman";
                    numberCell.Font.Size = 12;
                    numberCell.Font.Bold = 0;
                    numberCell.Font.Italic = 0;
                    numberCell.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    numberCell.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    numberCell.ParagraphFormat.SpaceAfter = 0f;
                    numberCell.ParagraphFormat.SpaceBefore = 0f;
                    numberCell.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                    table.Cell(row, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                }


                AddCustomParagraph(doc, "Судья:_______________________________", "Times New Roman",
                12, 0, 0, 0f, 2.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Тренер:_____________________________________________", "Times New Roman",
                12, 0, 0, 0f, 2.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Телефон тренера:_____________________________________________", "Times New Roman",
                12, 0, 0, 0f, 2.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Подписи", "Times New Roman",
                12, 0, 0, 0f, 1.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Руководитель______________(______________", "Times New Roman",
                12, 1, 1, 0f, 14.25f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Допущены:", "Times New Roman",
                12, 1, 1, 0f, 1.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpace1pt5, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "Главный судья______________\t\t\t\t\t\t Врач _______________________(______________", "Times New Roman",
                12, 1, 1, 0f, 1.5f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                AddCustomParagraph(doc, "М.П.", "Times New Roman",
                10, 0, 0, 0f, 17f, 1f, 0f, Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdLineSpacing.wdLineSpaceSingle, Word.WdUnderline.wdUnderlineNone);

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filePath = Path.Combine(desktopPath, "Заявочный лист на участие в соревнованиях.docx");
                doc.SaveAs2(filePath);

                MessageBox.Show("Документ создан на рабочем столе.");

                doc.Close(false);
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void AddEmptyLine(Word.Document doc)
        {
            object endOfDoc = "\\endofdoc";
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Word.Paragraph p = doc.Paragraphs.Add(ref range);
            p.Range.Text = "";
            p.Range.InsertParagraphAfter();
        }

        private void AddCustomParagraph(
            Word.Document doc,
            string text,
            string fontName,
            float fontSize,
            int bold,
            int italic,
            float leftIndentCm,
            float firstLineIndentCm,
            float spaceAfterPt,
            float spaceBeforePt,
            Word.WdParagraphAlignment alignment,
            Word.WdLineSpacing lineSpacingRule,
            Word.WdUnderline underlineStyle)
        {
            object endOfDoc = "\\endofdoc";
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Word.Paragraph p = doc.Paragraphs.Add(ref range);

            p.Range.Text = text;
            p.Range.Font.Name = fontName;
            p.Range.Font.Size = fontSize;
            p.Range.Font.Bold = bold;
            p.Range.Font.Italic = italic;
            p.Range.Font.Underline = underlineStyle;

            p.Range.ParagraphFormat.SpaceAfter = spaceAfterPt;
            p.Range.ParagraphFormat.SpaceBefore = spaceBeforePt;
            p.Range.ParagraphFormat.LeftIndent = doc.Application.CentimetersToPoints(leftIndentCm);
            p.Range.ParagraphFormat.FirstLineIndent = doc.Application.CentimetersToPoints(firstLineIndentCm);
            p.Range.ParagraphFormat.LineSpacingRule = lineSpacingRule;
            p.Alignment = alignment;

            p.Range.InsertParagraphAfter();
        }
    }
}