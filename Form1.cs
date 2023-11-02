using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class Form1 : Form
    {
        private Timer connectionStatusTimer;

        public Form1()
        {
            InitializeComponent();

            // Инициализируем и настраиваем таймер
            connectionStatusTimer = new Timer();
            connectionStatusTimer.Interval = 10000; // 20000 миллисекунд = 20 секунд
            connectionStatusTimer.Tick += new EventHandler(UpdateConnectionStatus);
        }

        private void UpdateConnectionStatus(object sender, EventArgs e)
        {
            if (CheckDatabaseConnection())
            {
                label1.Text = "Состояние подключения к базе данных: Подключено";
            }
            else
            {
                label1.Text = "Состояние подключения к базе данных: Отключено";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Запускаем таймер при загрузке формы
            connectionStatusTimer.Start();
        }

        private bool CheckDatabaseConnection()
        {
            DataBaseConnection dbConnection = new DataBaseConnection();
            NpgsqlConnection connection = dbConnection.GetConnection();
            try
            {
                connection.Open();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form5 form5 = new Form5();
            form5.Show();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
