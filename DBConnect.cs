using System;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore;
using Npgsql;
using ShiftReportApp1;

public class DataBaseConnection
{
    public string Host { get; set; }
    private string Port { get; set; }
    private string Database { get; set; }
    private string Username { get; set; }
    private string Password { get; set; }

    public DataBaseConnection()
    {
        // Конструктор может остаться пустым или содержать минимальную инициализацию.
        SetConnectionParameters();
    }

    public void SetConnectionParameters()
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
            }
            else
            {
                Console.WriteLine("Settings file not found.");
            }
        }
        catch (Exception ex)
        {
            // Обработка ошибки
            Console.WriteLine($"Error loading connection settings: {ex.Message}");
        }
    }

    public NpgsqlConnection GetConnection()
    {
        try
        {
            if (string.IsNullOrEmpty(Host))
            {
                // Вы можете обработать ситуацию, когда Host равен null или пуст
                throw new ArgumentException("Host can't be null or empty");
            }

            string connectionString = $"Host={Host};Port={Port};Database={Database};Username={Username};Password={Password}";
            return new NpgsqlConnection(connectionString);
        }
        catch (Exception ex)
        {
            // Обработка ошибки
            MessageBox.Show($"Ошибка подключения к SQL server: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            ProjectLogger.LogException("Ошибка подключения к SQL server", ex);
            return null;
        }
    }
}