using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Shoe_Organizer__Excel_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e) //Удаление по выделенному индексу
        {
            try
            {
                if (dataGridView1.Rows.Count > 1)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
                }
            }
            catch (Exception s)
            {

                MessageBox.Show(s.Message);
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e) //Удаление по запросу
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (Convert.ToString(dataGridView1[0, i].Value) == Convert.ToString(textBox1.Text))
                    {
                        dataGridView1.Rows.RemoveAt(i);
                    }
                }
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e) //Кнопка Увеличить стоимость
        {
            try
            {
                if (dataGridView1.Rows.Count > 1)
                {
                    if (dataGridView1.CurrentCell.ColumnIndex == 4 && textBox2.Text != "")
                    {
                        decimal salary_to_increase = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[4].Value);
                        salary_to_increase *= (Convert.ToDecimal(textBox2.Text) / 100) + 1; 
                        salary_to_increase = Math.Round(salary_to_increase, 2);
                        dataGridView1.CurrentRow.Cells[4].Value = salary_to_increase;
                    }
                    else
                    {
                        MessageBox.Show("Проверьте введенные данные");
                    }
                }
            }
            catch (Exception s)
            {

                MessageBox.Show(s.Message);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e) //Кнопка кол-во пар детской обуви
        {
            try
            {
                richTextBox1.Text = "";
                int kol = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    string m = Convert.ToString(dataGridView1[0, i].Value);
                    char[] line = m.ToCharArray();

                    for (int j = 0; j < line.Length; j++)
                    {
                        if (line[j] == 'Д' | line[j] == 'д')
                        {
                            kol += Convert.ToInt32(dataGridView1[3, i].Value);
                            break;
                        }
                    }
                }
                richTextBox1.Text += "Количество пар детской обуви:" + kol + "\n";
            }
            catch (Exception s)
            {

                MessageBox.Show(s.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e) //Кнопка "Стоимость всей обуви"
        {
            try
            {
                richTextBox1.Text = "";
                int sum = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    sum += (Convert.ToInt32(dataGridView1[4, i].Value) * Convert.ToInt32(dataGridView1[3, i].Value)); //Кол-во пар на стоимость одной пары
                }
                richTextBox1.Text += "Стоимость всей обуви:" + "\t" + sum + "\n";
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e) //Кнопка уменьшить
        {
            try
            {
                if (dataGridView1.Rows.Count > 1)
                {
                    if (dataGridView1.CurrentCell.ColumnIndex == 4 && textBox2.Text != "")
                    {
                        decimal salary_to_decrease = Convert.ToDecimal(dataGridView1.CurrentRow.Cells[4].Value);
                        salary_to_decrease *= 1 - (Convert.ToDecimal(textBox2.Text) / 100);
                        salary_to_decrease = Math.Round(salary_to_decrease, 2);


                        dataGridView1.CurrentRow.Cells[4].Value = salary_to_decrease;

                    }
                    else
                    {
                        MessageBox.Show("Проверьте введенные данные");
                    }
                }
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message);
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        struct DataParameter
        {
            public string FileName { get; set; }
        }

        DataParameter _inputParameter;

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) //БэкРабочий для записывания данных в файл
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workBookExcel = appExcel.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet worksheetExcel = null;
                string fileName = ((DataParameter)e.Argument).FileName;
                worksheetExcel = workBookExcel.Sheets[1];
                worksheetExcel = workBookExcel.ActiveSheet;
                //Копируем заголовки
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheetExcel.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                //Заполняем таблицу
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheetExcel.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                //Сохраняем и выбираем куда сохранять файл
                workBookExcel.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                appExcel.Visible = true;
                //appExcel.Quit();
            }
            catch (Exception s)
            {
                MessageBox.Show(s.Message);
            }
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e) //Кнопка Экспорт в меню Файл
        {
            if (backgroundWorker1.IsBusy)
                return;
            using (SaveFileDialog sfd = new SaveFileDialog() {Filter = "Excel Workbook|*.xls" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _inputParameter.FileName = sfd.FileName;
                    backgroundWorker1.RunWorkerAsync(_inputParameter);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int d = int.Parse(textBox3.Text);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string m = Convert.ToString(dataGridView1[0, i].Value);
                char[] line = m.ToCharArray();

                for (int j = 0; j < line.Length; j++)
                {
                    if (line[j] == 'М' | line[j] == 'м')
                    {
                        int b = int.Parse(Convert.ToString(dataGridView1[4, i].Value));

                        if (b < d)
                        {
                            dataGridView1.Rows.RemoveAt(i);

                        }

                    }
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            int d = int.Parse(textBox4.Text);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string m = Convert.ToString(dataGridView1[0, i].Value);
                char[] line = m.ToCharArray();

                for (int j = 0; j < line.Length; j++)
                {
                    if (line[j] == 'М' | line[j] == 'м')
                    {
                        int b = int.Parse(Convert.ToString(dataGridView1[4, i].Value));

                        if (b < d)
                        {
                            dataGridView1.Rows.RemoveAt(i);

                        }

                    }
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
