using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyCourseWork
{
    public partial class Form1 : Form
    {
        DataTable table = new DataTable();
        private double monthlyEarnings;
        private string selectedEmployeeName = string.Empty;

        public Form1()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            LoadEmployeeList();
            SetTextBoxesReadOnly(true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "Виберіть працівника із списку";
            label1.ForeColor = Color.Blue;
            label2.Text = "Посада:";
            label3.Text = "Середньоденна зарплата:";
            label4.Text = "Заробітна плата:";
            label5.Text = "Відпустка з";
            label6.Text = "по";
            label7.Text = "Кількість відпускних днів:";
            label8.Text = "Нараховано:";
            label8.ForeColor = Color.Blue;
            label9.Text = "Оберіть дати";
            label9.ForeColor = Color.Blue;
            label10.Text = "Інформація про працівника";
            label10.ForeColor = Color.Blue;
            button1.Text = "Нарахувати відпускні";
            button1.BackColor = Color.RoyalBlue;
            button1.ForeColor = Color.White;
        }

        private void SetTextBoxesReadOnly(bool readOnly)
        {
            textBox1.ReadOnly = readOnly;
            textBox2.ReadOnly = readOnly;
            textBox3.ReadOnly = readOnly;
            textBox4.ReadOnly = readOnly;
            textBox5.ReadOnly = readOnly;
        }

        private void LoadEmployeeList()
        {
            string filePath = @"C:\Users\victo\OneDrive\Документы\my\2-ий курс\OOP\MyCourseWork\MyCourseWork\bin\Debug\personal.xlsx";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    table = result.Tables[0];
                }
            }

            listBox1.DataSource = table;
            listBox1.ValueMember = table.Columns[0].ColumnName;
            listBox1.DisplayMember = table.Columns[0].ColumnName;
        }

        private int getDaysOffNumber(DateTime startDate, DateTime endDate)
        {
            int daysOff = 0;
            for (DateTime date = startDate; date < endDate; date = date.AddDays(1))
            {
                daysOff++;
            }
            return daysOff;
        }

        private void UpdateDaysOffNumber()
        {
            if (startDateTimePicker1.Value > endDateTimePicker2.Value)
            {
                MessageBox.Show("Дата початку не може бути пізнішою, ніж дата кінця, або співпадати з нею.");
                textBox3.Text = string.Empty;
                return;
            }

            int daysOffNumber = getDaysOffNumber(startDateTimePicker1.Value, endDateTimePicker2.Value);

            if (daysOffNumber > 24)
            {
                MessageBox.Show("Кількість відпусткових днів не може перевищувати 24.");
                textBox3.Text = string.Empty;
                return;
            }

            textBox3.Text = daysOffNumber.ToString();

        }

        private double CalculateTotalSalary(int daysOffNumber)
        {
            double averageEarningForDay = CalculateAverageEarningForDay();

            double totalSalary = averageEarningForDay * daysOffNumber;

            return totalSalary;
        }

        private double CalculateAverageEarningForDay()
        {
            double maxAverageEarningForDay = 3301.58;

            double averageEarningForDay = monthlyEarnings * 12 / 365;

            if (averageEarningForDay > maxAverageEarningForDay)
            {
                averageEarningForDay = maxAverageEarningForDay;
            }

            return Math.Round(averageEarningForDay);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = listBox1.SelectedIndex;

            if (selectedIndex != -1)
            {
                DataRowView selectedRow = (DataRowView)listBox1.SelectedItem;

                string fullName = selectedRow.Row[0].ToString();

                DataRow[] filteredRows = table.Select($"{table.Columns[0].ColumnName} = '{fullName}'");


                if (filteredRows.Length > 0)
                {
                    selectedEmployeeName = fullName;
                    monthlyEarnings = double.Parse(filteredRows[0][2].ToString());

                    textBox1.Text = filteredRows[0][1].ToString();
                    textBox2.Text = filteredRows[0][2].ToString();
                    textBox5.Text = CalculateAverageEarningForDay().ToString();

                    textBox4.Text = string.Empty;
                    startDateTimePicker1.Value = DateTime.Today;
                    endDateTimePicker2.Value = DateTime.Today;
                    textBox3.Text = string.Empty;
                }
            }
        }

        private void startDateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = string.Empty;
            UpdateDaysOffNumber();
        }

        private void endDateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = string.Empty;
            UpdateDaysOffNumber();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int daysOffNumber = getDaysOffNumber(startDateTimePicker1.Value, endDateTimePicker2.Value);

            if (!string.IsNullOrEmpty(selectedEmployeeName))
            {
                double totalSalary = CalculateTotalSalary(daysOffNumber);

                if (daysOffNumber > 0 && textBox3.Text != string.Empty)
                {
                    textBox4.Text = totalSalary.ToString();
                }
                else
                {
                    MessageBox.Show("Оберіть коректно дати початку та кінця відпустки");
                }
            }
        }
    }
}
