﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Exel_3._0
{
    public partial class MainMenu : Form
    {
        private const int col_d_row = 65536;
        private const int coordinates = 30;
        private const int width_size = 40;           //розміри клітинки
        private const int height_correct = 30;       //відступи по висоті
        private const int alphabet = 65;
        private const int number_of_letters = 26;
        private const int start_number = 2;         //кількість початкових стовпчиків/рядків
        public int col = 0;
        public string[,] massFormulas; //масив формул в парсер
        public MainMenu()
        {
            InitializeComponent();
            ChangeSize();
            for (int i = 0; i < start_number; i++)
            {
                AddColumn();
                AddRow();
            }
        }

        private void FillingMass()
        {
            massFormulas = new string[dataGridView1.ColumnCount, dataGridView1.RowCount];
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    if (dataGridView1[i, j].Value != null)
                        massFormulas[i, j] = dataGridView1[i, j].Value.ToString();
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        public string Formulas(int i, int j)
        {
            if (dataGridView1[i, j].Value != null)
            {
                string s = dataGridView1[0, 0].Value.ToString();
                return s;
            }
            else
                return "0";
        }

        private void AddColumn()
        {
            int temp = 0;
            dataGridView2.ColumnCount++;
            temp = dataGridView1.ColumnCount++;
            dataGridView1.Columns[temp].HeaderText = Convert.ToString(temp + 1);
            dataGridView1.Columns[temp].Width = width_size;
            dataGridView2.Columns[temp].HeaderText = Convert.ToString(temp + 1);
            dataGridView2.Columns[temp].Width = width_size;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        private void AddRow()
        {
            int temp = 0;
            if ((dataGridView1.RowCount == 0) && (dataGridView1.ColumnCount == 0))
            {
                AddColumn();
            }
            dataGridView2.RowCount++;
            temp = dataGridView1.RowCount++;
            if (temp < number_of_letters)
            {
                dataGridView1.Rows[temp].HeaderCell.Value = Convert.ToString((char)(temp + alphabet));
                dataGridView2.Rows[temp].HeaderCell.Value = Convert.ToString((char)(temp + alphabet));
            }
            else
                if (temp < (number_of_letters * (number_of_letters + 1)))
                {
                    int secondT = temp % number_of_letters;
                    int firstT = temp / number_of_letters - 1;
                    string s;
                    s = Convert.ToString((char)(firstT + alphabet));
                    s += Convert.ToString((char)(secondT + alphabet));
                    dataGridView1.Rows[temp].HeaderCell.Value = s;
                    dataGridView2.Rows[temp].HeaderCell.Value = s;
                }
                else
                    if (temp < (number_of_letters * (number_of_letters * (number_of_letters + 1) + 1)))
                    {

                        int thirdT = temp % number_of_letters;
                        int secondT = temp / number_of_letters - number_of_letters - 1;
                        int firstT = temp / number_of_letters / number_of_letters - 1;
                        string s;
                        s = Convert.ToString((char)(firstT + alphabet));
                        s += Convert.ToString((char)(secondT + alphabet));
                        s += Convert.ToString((char)(thirdT + alphabet));
                        dataGridView1.Rows[temp].HeaderCell.Value = s;
                        dataGridView2.Rows[temp].HeaderCell.Value = s;
                    }
                    else
                    {
                        dataGridView1.RowCount--; dataGridView2.RowCount--;
                    }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ChangeSide();
        }

        void ChangeSide()
        {
            for (int j = 0; j < dataGridView2.RowCount; j++)
            {
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    if (dataGridView1[i, j].Value == null)
                    {
                        dataGridView1[i, j].Value = dataGridView2[i, j].Value;
                        dataGridView1[i, j].Value = "0";
                    }
                }
            }
            Recalculation();
            dataGridView2.Width = dataGridView1.Width;
            dataGridView2.Height = dataGridView1.Height;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
        }

        private void Recalculation()
        {
            FillingMass();
            for (int j = 0; j < dataGridView1.RowCount; j++)
            {
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1[i, j].Value != null)
                    {
                        int ColumnCount = dataGridView1.ColumnCount;
                        int RowCount = dataGridView2.RowCount;
                        Parser a = new Parser(massFormulas, ColumnCount, RowCount);
                        string s = dataGridView1[i, j].Value.ToString();
                        double rez = a.Evaluate(s);
                        dataGridView2[i, j].Value = Convert.ToString(rez);
                    }
                }
            }
        }
        private void MainMenu_SizeChanged(object sender, EventArgs e)
        {
            ChangeSize();
        }
        private void ChangeSize()
        {
            int delta_width = 0;
            int delta_height = 0;
            dataGridView1.Location = new Point(coordinates, coordinates);
            dataGridView2.Location = new Point(coordinates, coordinates);
            delta_width = -button1.Width - width_size / 2;
            delta_height = -button1.Height - button2.Height - height_correct;
            button1.Location = new Point(this.Size.Width + delta_width, this.Size.Height + delta_height);
            delta_width -= button4.Width;
            button4.Location = new Point(this.Size.Width + delta_width, this.Size.Height + delta_height);
            delta_width -= button2.Width;
            button2.Location = new Point(this.Size.Width + delta_width, this.Size.Height + delta_height);
            delta_width -= button5.Width;
            button5.Location = new Point(this.Size.Width + delta_width, this.Size.Height + delta_height);
            delta_width -= button3.Width;
            button3.Location = new Point(this.Size.Width + delta_width, this.Size.Height + delta_height);
            dataGridView1.Width = this.Size.Width - width_size - coordinates;
            dataGridView1.Height = this.Size.Height - 3 * height_correct - coordinates;
            dataGridView2.Width = this.Size.Width - width_size - coordinates;
            dataGridView2.Height = this.Size.Height - 3 * height_correct - coordinates;
        }
        private void Ntab(int ColumnC, int RowC, string[,] mass)
        {
            dataGridView1.ColumnCount = 0;
            dataGridView2.ColumnCount = 0;
            for (int i = 0; i < ColumnC; i++)
            {
                AddColumn();
            }
            dataGridView1.RowCount = 0;
            dataGridView2.RowCount = 0;
            for (int i = 0; i < RowC; i++)
            {
                AddRow();
            }
            for (int i = 0; i < ColumnC; i++)
            {
                for (int j = 0; j < RowC; j++)
                {
                    dataGridView1[i, j].Value = mass[i, j];
                }
            }
            ChangeSide();
        }
        private void проПрограмуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Created by Yaroshevska Nadia");
        }
        private void відкритиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = "c:\\";
            ofd.Filter = "Exel 3.0 files (*.exel3)|*.exel3";
            ofd.FilterIndex = 2;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.Cancel)
                return;
            StreamReader sr = new StreamReader(ofd.FileName);
            int ColumnCount_ = 0;
            int RowCount_ = 0;
            string[,] mass;
            if (sr.Peek() >= 0) ColumnCount_ = Convert.ToInt32(sr.ReadLine());
            else
            {
                Show a = new Show();
                a.FILEFAIL();
                return;
            }
            if (sr.Peek() >= 0) RowCount_ = Convert.ToInt32(sr.ReadLine());
            else
            {
                Show a = new Show();
                a.FILEFAIL();
                return;
            }
            if (RowCount_ * ColumnCount_ > col_d_row)
            {
                Show a = new Show();
                a.FILEFAIL();
                return;
            }
            mass = new string[ColumnCount_, RowCount_];
            for (int i = 0; i < ColumnCount_; i++)
            {
                for (int j = 0; j < RowCount_; j++)
                {
                    if (sr.Peek() >= 0)
                    {
                        mass[i, j] = sr.ReadLine();
                        if (mass[i, j] == "0") mass[i, j] = "";
                    }
                    else
                    {
                        Show a = new Show();
                        a.FILEFAIL();
                        return;
                    }
                }
            }
            Ntab(ColumnCount_, RowCount_, mass);
            sr.Close();
        }
        private void зберегтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Exel 3.0 files (*.exel3)|*.exel3";
                sfd.FilterIndex = 2;
                sfd.RestoreDirectory = true;
                if (sfd.ShowDialog() == DialogResult.Cancel)
                    return;
                StreamWriter sw = new StreamWriter(sfd.FileName);
                sw.WriteLine(dataGridView1.ColumnCount);
                sw.WriteLine(dataGridView1.RowCount);
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                    {
                        if (dataGridView1[i, j].Value != null)
                            sw.WriteLine(Convert.ToString(dataGridView1[i, j].Value));
                    }
                }
                MessageBox.Show("Файл успішно збережено");
                sw.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Виникла помилка"+ "\n" + "Файл не збережено");
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DeleteCount();
        }
        private void DeleteCount()
        {
            if (dataGridView1.ColumnCount > 0)
            {
                int k = dataGridView2.CurrentCell.ColumnIndex;
                for (int i = dataGridView2.ColumnCount - 1; i > k; i--)
                {
                    dataGridView1.Columns[i].HeaderText = dataGridView1.Columns[i - 1].HeaderText;
                    dataGridView2.Columns[i].HeaderText = dataGridView2.Columns[i - 1].HeaderText;
                }
                dataGridView1.Columns.RemoveAt(k);
                dataGridView2.Columns.RemoveAt(k);
            }
            ChangeSide();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DeleteRow();
        }
        private void DeleteRow()
        {
            if (dataGridView1.RowCount > 0)
            {
                int k = dataGridView2.CurrentRow.Index;
                for (int i = dataGridView2.RowCount - 1; i > k; i--)
                {
                    dataGridView1.Rows[i].HeaderCell.Value = dataGridView1.Rows[i - 1].HeaderCell.Value;
                    dataGridView2.Rows[i].HeaderCell.Value = dataGridView2.Rows[i - 1].HeaderCell.Value;
                }
                dataGridView1.Rows.RemoveAt(k);
                dataGridView2.Rows.RemoveAt(k);

            }
            if (dataGridView1.RowCount == 0)
            {
                dataGridView1.ColumnCount = 0;
                dataGridView2.ColumnCount = 0;
            }
            ChangeSide();
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex != -1)
            {
                var a = dataGridView1[e.ColumnIndex, e.RowIndex].Value;
                if (a != null)
                    textBox1.Text = a.ToString();
            }
        }
        private void dataGridView2_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataGridView2[e.ColumnIndex, e.RowIndex].Value = dataGridView1[e.ColumnIndex, e.RowIndex].Value;
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1[e.ColumnIndex, e.RowIndex].Value = dataGridView2[e.ColumnIndex, e.RowIndex].Value;
            ChangeSide();
        }
    }
    public partial class Show : Form
    {
        public void INCORRECT()
        {
            MessageBox.Show("Некоректне посилання на клітинку");
        }
        public void FILEFAIL()
        {
            MessageBox.Show("Помилка у читаннi файла");
        }
        public void CYCLE()
        {
            MessageBox.Show("Зациклення");
        }
        public void DIVBYZERO()
        {
            MessageBox.Show("Дiлення на 0");
        }
    }
}
