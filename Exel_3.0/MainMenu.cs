using System;
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
        private const int start_number = 5;         //кількість початкових стовпчиків/рядків
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
            massFormulas = new string[dgFormula.ColumnCount, dgFormula.RowCount];
            for (int i = 0; i < dgFormula.ColumnCount; i++)
            {
                for (int j = 0; j < dgFormula.RowCount; j++)
                {
                    if (dgFormula[i, j].Value != null)
                        massFormulas[i, j] = dgFormula[i, j].Value.ToString();
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        public string Formulas(int i, int j)
        {
            if (dgFormula[i, j].Value != null)
            {
                string s = dgFormula[0, 0].Value.ToString();
                return s;
            }
            else
                return "0";
        }

        private void AddColumn()
        {
            int temp = 0;
            dgValue.ColumnCount++;
            temp = dgFormula.ColumnCount++;
            dgFormula.Columns[temp].HeaderText = Convert.ToString(temp + 1);
            dgFormula.Columns[temp].Width = width_size;
            dgValue.Columns[temp].HeaderText = Convert.ToString(temp + 1);
            dgValue.Columns[temp].Width = width_size;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        private void AddRow()
        {
            int temp = 0;
            if ((dgFormula.RowCount == 0) && (dgFormula.ColumnCount == 0))
            {
                AddColumn();
            }
            dgValue.RowCount++;
            temp = dgFormula.RowCount++;
            if (temp < number_of_letters)
            {
                dgFormula.Rows[temp].HeaderCell.Value = Convert.ToString((char)(temp + alphabet));
                dgValue.Rows[temp].HeaderCell.Value = Convert.ToString((char)(temp + alphabet));
            }
            else
                if (temp < (number_of_letters * (number_of_letters + 1)))
                {
                    int secondT = temp % number_of_letters;
                    int firstT = temp / number_of_letters - 1;
                    string s;
                    s = Convert.ToString((char)(firstT + alphabet));
                    s += Convert.ToString((char)(secondT + alphabet));
                    dgFormula.Rows[temp].HeaderCell.Value = s;
                    dgValue.Rows[temp].HeaderCell.Value = s;
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
                        dgFormula.Rows[temp].HeaderCell.Value = s;
                        dgValue.Rows[temp].HeaderCell.Value = s;
                    }
                    else
                    {
                        dgFormula.RowCount--; dgValue.RowCount--;
                    }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ChangeSide();
        }

        void ChangeSide()
        {
            for (int j = 0; j < dgValue.RowCount; j++)
            {
                for (int i = 0; i < dgValue.ColumnCount; i++)
                {
                    if (dgFormula[i, j].Value == null)
                    {
                        dgFormula[i, j].Value = dgValue[i, j].Value;
                        //dgFormula[i, j].Value = "0";
                    }
                }
            }
            Recalculation();
            dgValue.Width = dgFormula.Width;
            dgValue.Height = dgFormula.Height;
            dgFormula.Visible = false;
            dgValue.Visible = true;
        }

        private void Recalculation()
        {
            FillingMass();
            for (int j = 0; j < dgFormula.RowCount; j++)
            {
                for (int i = 0; i < dgFormula.ColumnCount; i++)
                {
                    if (dgFormula[i, j].Value != null)
                    {
                        int ColumnCount = dgFormula.ColumnCount;
                        int RowCount = dgValue.RowCount;
                        Parser a = new Parser(massFormulas, ColumnCount, RowCount);
                        string s = dgFormula[i, j].Value.ToString();
                        string rez = a.StartEvaluate(s);
                        dgValue[i, j].Value = rez;
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
            dgFormula.Location = new Point(coordinates, coordinates);
            dgValue.Location = new Point(coordinates, coordinates);
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
            dgFormula.Width = this.Size.Width - width_size - coordinates;
            dgFormula.Height = this.Size.Height - 3 * height_correct - coordinates;
            dgValue.Width = this.Size.Width - width_size - coordinates;
            dgValue.Height = this.Size.Height - 3 * height_correct - coordinates;
        }
        private void Ntab(int ColumnC, int RowC, string[,] mass)
        {
            dgFormula.ColumnCount = 0;
            dgValue.ColumnCount = 0;
            for (int i = 0; i < ColumnC; i++)
            {
                AddColumn();
            }
            dgFormula.RowCount = 0;
            dgValue.RowCount = 0;
            for (int i = 0; i < RowC; i++)
            {
                AddRow();
            }
            for (int i = 0; i < ColumnC; i++)
            {
                for (int j = 0; j < RowC; j++)
                {
                    dgFormula[i, j].Value = mass[i, j];
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
                sw.WriteLine(dgFormula.ColumnCount);
                sw.WriteLine(dgFormula.RowCount);
                for (int i = 0; i < dgFormula.ColumnCount; i++)
                {
                    for (int j = 0; j < dgFormula.RowCount; j++)
                    {
                        if (dgFormula[i, j].Value != null)
                            sw.WriteLine(Convert.ToString(dgFormula[i, j].Value));
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
            if (dgFormula.ColumnCount > 0)
            {
                int k = dgValue.CurrentCell.ColumnIndex;
                for (int i = dgValue.ColumnCount - 1; i > k; i--)
                {
                    dgFormula.Columns[i].HeaderText = dgFormula.Columns[i - 1].HeaderText;
                    dgValue.Columns[i].HeaderText = dgValue.Columns[i - 1].HeaderText;
                }
                dgFormula.Columns.RemoveAt(k);
                dgValue.Columns.RemoveAt(k);
            }
            ChangeSide();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DeleteRow();
        }
        private void DeleteRow()
        {
            if (dgFormula.RowCount > 0)
            {
                int k = dgValue.CurrentRow.Index;
                for (int i = dgValue.RowCount - 1; i > k; i--)
                {
                    dgFormula.Rows[i].HeaderCell.Value = dgFormula.Rows[i - 1].HeaderCell.Value;
                    dgValue.Rows[i].HeaderCell.Value = dgValue.Rows[i - 1].HeaderCell.Value;
                }
                dgFormula.Rows.RemoveAt(k);
                dgValue.Rows.RemoveAt(k);

            }
            if (dgFormula.RowCount == 0)
            {
                dgFormula.ColumnCount = 0;
                dgValue.ColumnCount = 0;
            }
            ChangeSide();
        }
        private void dgValue_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != -1 && e.RowIndex != -1)
            {
                var a = dgFormula[e.ColumnIndex, e.RowIndex].Value;
                if (a != null)
                    textBox1.Text = a.ToString();
            }
        }
        private void dgValue_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dgValue[e.ColumnIndex, e.RowIndex].Value = dgFormula[e.ColumnIndex, e.RowIndex].Value;
        }

        private void dgValue_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgFormula[e.ColumnIndex, e.RowIndex].Value = dgValue[e.ColumnIndex, e.RowIndex].Value;
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
