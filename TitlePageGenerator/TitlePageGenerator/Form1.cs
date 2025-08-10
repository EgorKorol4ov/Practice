using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace TitlePageGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            string path = Application.StartupPath + "\\settings.txt";
            var settings = new List<ParagraphSetting>();

            var vypolnil = new ParagraphSetting
            {
                Text = "Выполнил: ст. гр. " + textBox_Group.Text + ", " + textBox_FIO.Text,
                Font = comboBox_Font.Text,
                FontSize = (float)numericUpDown_Size.Value,
                FontStyle = comboBox_Style.Text,
                LeftIndent = (float)numericUpDown_IndentVypolnil.Value
            };

            var proveril = new ParagraphSetting
            {
                Text = "Проверил: " + textBox_Proveril.Text,
                Font = comboBox_Font.Text,
                FontSize = (float)numericUpDown_Size.Value,
                FontStyle = comboBox_Style.Text,
                LeftIndent = (float)numericUpDown_IndentProveril.Value
            };

            settings.Add(vypolnil);
            settings.Add(proveril);

            File.WriteAllLines(path, settings.Select(s => s.ToString()));
            MessageBox.Show("Параметры сохранены в settings.txt");
        }

        private void button_CreateWord_Click(object sender, EventArgs e)
        {
            string path = Application.StartupPath + "\\settings.txt";

            if (!File.Exists(path))
            {
                MessageBox.Show("Файл settings.txt не найден! Сначала сохраните данные.");
                return;
            }

            var lines = File.ReadAllLines(path);
            var settings = lines.Select(l => ParagraphSetting.FromString(l)).Where(s => s != null).ToList();

            if (settings.Count == 0)
            {
                MessageBox.Show("Файл settings.txt пуст. Сначала сохраните данные.");
                return;
            }

            object missing = Missing.Value;
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Add();

            Word.Style normalStyle = doc.Styles["Обычный"];
            normalStyle.ParagraphFormat.LeftIndent = 0f;
            normalStyle.ParagraphFormat.SpaceAfter = 0f;
            normalStyle.ParagraphFormat.SpaceBefore = 0f;

            AddParagraphStyled(doc, "министерство транспорта российской федерации", 14, true, true);
            AddParagraphStyled(doc, "федеральное государственное автономное образовательное учреждение высшего образования", 14, true, false);
            AddParagraphStyled(doc, "«российский университет транспорта» (рут (миит))", 14, true, true);
            AddParagraphStyled(doc, "институт транспортной техники и систем управления", 14, true, true);

            Word.Range range = doc.Content.Paragraphs.Last.Range;
            Word.Table table = doc.Tables.Add(range, 1, 1, ref missing, ref missing);
            table.Borders.Enable = 0;
            table.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            table.Borders[Word.WdBorderType.wdBorderBottom].Color = Word.WdColor.wdColorBlack;
            table.Rows[1].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            table.Rows[1].Height = wordApp.CentimetersToPoints(0.5f);
            table.Cell(1, 1).Range.Text = " ";

            AddParagraphStyled(doc, "Кафедра «Управление и защита информации»", 14, false, false);

            AddParagraphStyled(doc, comboBox_DocType.Text, 17, false, false);
            AddParagraphStyled(doc, comboBox_TypeWork.Text, 14, false, false);
            AddParagraphStyled(doc, "Задание " + comboBox_WorkNumber.Text, 14, false, true);
            AddParagraphStyled(doc, "по дисциплине", 14, false, false);
            AddParagraphStyled(doc, textBox_Subject.Text, 14, false, true);
            AddParagraphStyled(doc, textBox_Topic.Text, 14, false, false);
            AddParagraphStyled(doc, "Вариант № " + comboBox_Variant.Text, 14, false, true);

            AddParagraphStyled(doc, "", 14, false, false);

            foreach (var s in settings)
            {
                Word.Paragraph para = doc.Content.Paragraphs.Add(ref missing);
                para.Range.Text = s.Text;
                para.Range.Font.Name = s.Font;
                para.Range.Font.Size = s.FontSize;
                para.Range.Font.Bold = s.FontStyle.Contains("Bold") ? 1 : 0;
                para.Range.Font.Italic = s.FontStyle.Contains("Italic") ? 1 : 0;
                para.Range.Font.Underline = s.FontStyle.Contains("Underline") ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone;
                para.LeftIndent = wordApp.CentimetersToPoints(s.LeftIndent);
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.InsertParagraphAfter();
            }

            for (int i = 0; i < 12; i++)
                AddParagraphStyled(doc, "", 14, false, false);

            Word.Range endRange = doc.Content;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            Word.Paragraph moscowParagraph = doc.Paragraphs.Add(endRange);
            moscowParagraph.Range.Text = "Москва " + DateTime.Now.Year + " г.";
            moscowParagraph.Range.Font.Name = "Times New Roman";
            moscowParagraph.Range.Font.Size = 14;
            moscowParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            moscowParagraph.LeftIndent = 0;

            string fileName = Application.StartupPath + "\\Титульный_лист_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".docx";
            doc.SaveAs2(fileName);

            MessageBox.Show("Документ создан:\n" + fileName);
        }

        private void AddParagraphStyled(Word.Document doc, string text, int fontSize, bool allCaps, bool bold)
        {
            object missing = Missing.Value;
            Word.Paragraph para = doc.Content.Paragraphs.Add(ref missing);
            para.Range.Text = text;
            para.Range.Font.Name = "Times New Roman";
            para.Range.Font.Size = fontSize;
            para.Range.Font.AllCaps = allCaps ? 1 : 0;
            para.Range.Font.Bold = bold ? 1 : 0;
            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            para.Range.InsertParagraphAfter();
        }
    }
}