using System.Windows.Forms;
using System;

namespace TitlePageGenerator
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.textBox_FIO = new System.Windows.Forms.TextBox();
            this.textBox_Group = new System.Windows.Forms.TextBox();
            this.textBox_Proveril = new System.Windows.Forms.TextBox();
            this.textBox_Subject = new System.Windows.Forms.TextBox();
            this.textBox_Topic = new System.Windows.Forms.TextBox();

            this.comboBox_DocType = new System.Windows.Forms.ComboBox();
            this.comboBox_WorkNumber = new System.Windows.Forms.ComboBox();
            this.comboBox_Var = new System.Windows.Forms.ComboBox();
            this.comboBox_Font = new System.Windows.Forms.ComboBox();
            this.comboBox_Style = new System.Windows.Forms.ComboBox();
            this.comboBox_TypeWork = new System.Windows.Forms.ComboBox();

            this.numericUpDown_Size = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown_IndentVypolnil = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown_IndentProveril = new System.Windows.Forms.NumericUpDown();

            this.button_Save = new System.Windows.Forms.Button();
            this.button_CreateWord = new System.Windows.Forms.Button();

            System.Windows.Forms.Label label_Font = new System.Windows.Forms.Label();
            label_Font.Location = new System.Drawing.Point(350, 170);
            label_Font.Text = "Шрифт (Выполнил/Проверил)";
            label_Font.AutoSize = true;

            System.Windows.Forms.Label label_Style = new System.Windows.Forms.Label();
            label_Style.Location = new System.Drawing.Point(350, 220);
            label_Style.Text = "Стиль (Выполнил/Проверил)";
            label_Style.AutoSize = true;

            System.Windows.Forms.Label label_Size = new System.Windows.Forms.Label();
            label_Size.Location = new System.Drawing.Point(350, 270);
            label_Size.Text = "Размер шрифта (Выполнил/Проверил)";
            label_Size.AutoSize = true;

            System.Windows.Forms.Label label_IndentVypolnil = new System.Windows.Forms.Label();
            label_IndentVypolnil.Location = new System.Drawing.Point(20, 220);
            label_IndentVypolnil.Text = "Отступ (Выполнил)";
            label_IndentVypolnil.AutoSize = true;

            System.Windows.Forms.Label label_IndentProveril = new System.Windows.Forms.Label();
            label_IndentProveril.Location = new System.Drawing.Point(150, 220);
            label_IndentProveril.Text = "Отступ (Проверил)";
            label_IndentProveril.AutoSize = true;

            System.Windows.Forms.Label label_DocType = new System.Windows.Forms.Label();
            label_DocType.Location = new System.Drawing.Point(350, 20);
            label_DocType.Text = "Тип документа";
            label_DocType.AutoSize = true;

            System.Windows.Forms.Label label_Var = new System.Windows.Forms.Label();
            label_Var.Location = new System.Drawing.Point(350, 70);
            label_Var.Text = "Номер задания";
            label_Var.AutoSize = true;

            System.Windows.Forms.Label label_TypeWork = new System.Windows.Forms.Label();
            label_TypeWork.Location = new System.Drawing.Point(350, 120);
            label_TypeWork.Text = "Вид работы";
            label_TypeWork.AutoSize = true;

            this.textBox_FIO.Location = new System.Drawing.Point(20, 20);
            this.textBox_FIO.Width = 300;
            this.textBox_FIO.Text = "ФИО";

            this.textBox_Group.Location = new System.Drawing.Point(20, 50);
            this.textBox_Group.Width = 300;
            this.textBox_Group.Text = "Группа";

            this.textBox_Proveril.Location = new System.Drawing.Point(20, 80);
            this.textBox_Proveril.Width = 300;
            this.textBox_Proveril.Text = "Проверил";

            this.textBox_Subject.Location = new System.Drawing.Point(20, 110);
            this.textBox_Subject.Width = 300;
            this.textBox_Subject.Text = "Дисциплина";

            this.textBox_Topic.Location = new System.Drawing.Point(20, 140);
            this.textBox_Topic.Width = 300;
            this.textBox_Topic.Text = "Тема";

            this.comboBox_DocType.Location = new System.Drawing.Point(350, 40);
            this.comboBox_DocType.Width = 150;
            this.comboBox_DocType.Items.AddRange(new string[]
            {
                "Отчёт", "Реферат", "Эссе", "Курсовой проект", "Курсовая работа", "Доклад", "Домашнее задание"
            });

            this.comboBox_TypeWork.Location = new System.Drawing.Point(350, 140);
            this.comboBox_TypeWork.Width = 150;
            this.comboBox_TypeWork.Items.AddRange(new string[]
            {
                "Лабораторная работа", "Практическая работа", "Индивидуальное задание", 
                "Учебная практика", "Производственная практика", "Преддипломная практика"
            });

            this.comboBox_WorkNumber.Location = new System.Drawing.Point(350, 90);
            this.comboBox_WorkNumber.Width = 150;
            for (int i = 1; i <= 10; i++) this.comboBox_WorkNumber.Items.Add(i.ToString());

            this.comboBox_Var.Location = new System.Drawing.Point(20, 170);
            this.comboBox_Var.Width = 150;
            this.comboBox_Var.Text = "Номер варианта";
            for (int i = 1; i <= 30; i++) this.comboBox_Var.Items.Add(i.ToString());

            this.comboBox_Font.Location = new System.Drawing.Point(350, 190);
            this.comboBox_Font.Width = 150;
            this.comboBox_Font.Items.AddRange(new string[] { "Times New Roman", "Calibri", "Arial" });

            this.comboBox_Style.Location = new System.Drawing.Point(350, 240);
            this.comboBox_Style.Width = 150;
            this.comboBox_Style.Items.AddRange(new string[] { "Normal", "Bold", "Italic", "Underline" });

            this.numericUpDown_Size.Location = new System.Drawing.Point(350, 290);
            this.numericUpDown_Size.Minimum = 8;
            this.numericUpDown_Size.Maximum = 36;
            this.numericUpDown_Size.Value = 14;

            this.numericUpDown_IndentVypolnil.Location = new System.Drawing.Point(20, 240);
            this.numericUpDown_IndentVypolnil.Maximum = 15;
            this.numericUpDown_IndentVypolnil.Value = 10;

            this.numericUpDown_IndentProveril.Location = new System.Drawing.Point(150, 240);
            this.numericUpDown_IndentProveril.Maximum = 15;
            this.numericUpDown_IndentProveril.Value = 10;

            this.button_Save.Location = new System.Drawing.Point(20, 320);
            this.button_Save.Text = "Сохранить";
            this.button_Save.Click += new System.EventHandler(this.button_Save_Click);

            this.button_CreateWord.Location = new System.Drawing.Point(150, 320);
            this.button_CreateWord.Text = "Создать Word";
            this.button_CreateWord.Click += new System.EventHandler(this.button_CreateWord_Click);

            this.ClientSize = new System.Drawing.Size(550, 370);
            this.Controls.AddRange(new System.Windows.Forms.Control[]
            {
                this.textBox_FIO, this.textBox_Group, this.textBox_Proveril, 
                this.textBox_Subject, this.textBox_Topic, this.comboBox_DocType, this.comboBox_WorkNumber,
                this.comboBox_Font, this.comboBox_Style, this.comboBox_TypeWork, this.numericUpDown_Size,
                this.numericUpDown_IndentVypolnil, this.numericUpDown_IndentProveril,
                this.button_Save, this.button_CreateWord, this.comboBox_Var,
                label_Font, label_Style, label_Size, label_IndentVypolnil, label_IndentProveril, label_DocType, label_Var, label_TypeWork
            });

            this.Text = $"Задание №01 выполнил: Королёв Е.С.; Номер варианта: 13; Дата: {DateTime.Now:dd/MM/yyyy}";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.TextBox textBox_FIO;
        private System.Windows.Forms.TextBox textBox_Group;
        private System.Windows.Forms.TextBox textBox_Proveril;
        private System.Windows.Forms.TextBox textBox_Subject;
        private System.Windows.Forms.TextBox textBox_Topic;
        private System.Windows.Forms.ComboBox comboBox_DocType;
        private System.Windows.Forms.ComboBox comboBox_Var;
        private System.Windows.Forms.ComboBox comboBox_WorkNumber;
        private System.Windows.Forms.ComboBox comboBox_Font;
        private System.Windows.Forms.ComboBox comboBox_Style;
        private ComboBox comboBox_TypeWork;
        private System.Windows.Forms.NumericUpDown numericUpDown_Size;
        private System.Windows.Forms.NumericUpDown numericUpDown_IndentVypolnil;
        private System.Windows.Forms.NumericUpDown numericUpDown_IndentProveril;
        private System.Windows.Forms.Button button_Save;
        private System.Windows.Forms.Button button_CreateWord;
    }
}