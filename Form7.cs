using Npgsql;
using System;
using System.Data;
using System.Reflection.Emit;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class Form7 : Form
    {
        public Form7(DateTime varDate1, DateTime varDate2, int queryNumber, int numShift1, int numShift2, int numShift3,
            int numShift4, string typeStops1, string typeStops2, string typeStops3, string typeStops4)
        {
            try
            {
                InitializeComponent();
                label1.Text = "Даты с  " + varDate1.ToString() + "  по  " + varDate2.ToString();
                FillDataGridView(queryNumber, varDate1, varDate2, numShift1, numShift2, numShift3, numShift4,
                    typeStops1, typeStops2, typeStops3, typeStops4);
            }
            catch (Exception ex)
            {
                // Обработка ошибки
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в FillDataGridView", ex);
            }
        }

        private void FillDataGridView(int queryNumber, DateTime varDate1, DateTime varDate2, int numShift1, int numShift2,
            int numShift3, int numShift4, string typeStops1, string typeStops2, string typeStops3, string typeStops4)
        {
            string query = Names.Request(queryNumber);

            using (NpgsqlConnection connection = new DataBaseConnection().GetConnection())
            {
                connection.Open();
                using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
                {
                    if (varDate1 > varDate2)
                    {
                        DateTime tempDate = varDate1;
                        varDate1 = varDate2;
                        varDate2 = tempDate;
                    }

                    cmd.Parameters.AddWithValue("@varDate1", varDate1);
                    cmd.Parameters.AddWithValue("@varDate2", varDate2);
                    cmd.Parameters.AddWithValue("@numShift1", numShift1);
                    cmd.Parameters.AddWithValue("@numShift2", numShift2);
                    cmd.Parameters.AddWithValue("@numShift3", numShift3);
                    cmd.Parameters.AddWithValue("@numShift4", numShift4);
                    cmd.Parameters.AddWithValue("@typeStops1", typeStops1);
                    cmd.Parameters.AddWithValue("@typeStops2", typeStops2);
                    cmd.Parameters.AddWithValue("@typeStops3", typeStops3);
                    cmd.Parameters.AddWithValue("@typeStops4", typeStops4);

                    using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(cmd))
                    {
                        DataTable dataTable = new DataTable();
                        adapter.Fill(dataTable);
                        dataGridView1.DataSource = dataTable;
                    }
                }
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
