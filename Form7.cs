using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection.Emit;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class Form7 : Form
    {
        public int GetMethod { get; set; }
        public Form7(int getMethod, DateTime varDate1, DateTime varDate2, List<int>shiftDays, List<int> shifts, List<string> stopCategoryes)
        {
            InitializeComponent();
            GetMethod = getMethod;
            label1.Text = "Даты с  " + varDate1.ToString() + "  по  " + varDate2.ToString();
            // FillDataGridView(queryNumber, varDate1, varDate2, numShift1, numShift2, numShift3, numShift4,
            //    typeStops1, typeStops2, typeStops3, typeStops4);
            FillTable(getMethod, varDate1, varDate2, shiftDays, shifts, stopCategoryes);

        }

        private void FillTable(int getMethod, DateTime varDate1, DateTime varDate2, List<int> shiftDays, List<int> shifts, List<string> stopCategoryes)
        {
            LINQRequest newReport = new LINQRequest();
            DataTable dataTable = newReport.ExtractProduct(getMethod, varDate1, varDate2, shiftDays, shifts, stopCategoryes);
            dataTable.Rows.Add();
            dataGridView1.DataSource = dataTable;
            
            if (getMethod >= 0 && getMethod <= 5)
            {
                dataGridView1.CellFormatting += dataGridView1_CellFormatting1;
            }
        }

        private void dataGridView1_CellFormatting1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            List<int> max = new List<int> { };
            if (GetMethod == 4)
            {
                max = new List<int> { 2, 3, 4};
            }
            else if (GetMethod == 5)
            {
                max = new List<int> { 3, 4, 5 };
            }
            else if (GetMethod >= 0 && GetMethod <= 3)
            {
                max =  new List<int> { 7, 8, 9, 11, 12, 13, 14, 15};
            }
            foreach (var col in max)
            {
                decimal sum = 0;

                for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count - 1; rowIndex++)
                {
                    if (dataGridView1[col, rowIndex].Value != null &&
                        decimal.TryParse(dataGridView1[col, rowIndex].Value.ToString(), out decimal cellValue))
                    {
                        sum += cellValue;
                    }
                }

                // Отображение суммы в нужной ячейке
                dataGridView1[col, dataGridView1.Rows.Count -1 ].Value = sum;
            }
        }

        private void Form7_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
