using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace ShiftReportApp1
{
    public partial class Form9 : Form
    {
        private string Host { get; set; }
        private string Port { get; set; }
        private string Database { get; set; }
        private string Username { get; set; }
        private string Password { get; set; }

        private void SaveSettings()
        {
            try
            {
                string filePath = "CS.tx_";
                if (File.Exists(filePath))
                {
                    string[] settings = File.ReadAllText(filePath).Split(',');

                    if (settings.Length == 5)
                    {
                        Host = settings[0];
                        Port = settings[1];
                        Database = settings[2];
                        Username = settings[3];
                        Password = settings[4];
                    }
                    else
                    {
                        Console.WriteLine("Invalid format in the settings file.");
                    }
                    if (textBox1.Text != "") Host = textBox1.Text;
                    if (textBox2.Text != "") Port = textBox2.Text;
                    if (textBox5.Text != "") Database = textBox5.Text;
                    if (textBox4.Text != "") Username = textBox4.Text;
                    if (textBox3.Text != "") Password = textBox3.Text;

                    string newSettings = $"{Host},{Port},{Database},{Username},{Password}";

                    File.WriteAllText(filePath, newSettings);
                }
                else
                {
                    Console.WriteLine("Settings file not found.");
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибки
                Console.WriteLine($"Error rewrite connection settings: {ex.Message}");
            }

        }

        public Form9()
        {
            InitializeComponent();
            textBox3.Enabled = textBox4.Enabled = textBox5.Enabled = false;
        }

        private void Form9_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Enabled = textBox4.Enabled = textBox5.Enabled = checkBox1.Checked ? true : false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                SaveSettings();
                MessageBox.Show("Данные сохранены!");

                textBox1.Text = textBox2.Text = textBox3.Text = textBox4.Text = textBox5.Text = "";
            }
            catch
            {
                MessageBox.Show("Данные не сохранены. Проверьте введенные данные.");
                return;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form9 form9 = Application.OpenForms.OfType<Form9>().FirstOrDefault();
            form9.Close();
        }
    }
}
