using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.EntityFrameworkCore;
using Npgsql;
using ShiftReportApp1;

public class DataBaseConnection
{

    private string Host { get; }
    private int Port { get; }
    private string Database { get; }
    private string Username { get; }
    private string Password { get; }

    public DataBaseConnection(string host = "localhost", int port = 5432,
        string database = "reportdb", string username = "UserDB", string password = "UserOperate!1")
    {
        Host = host;
        Port = port;
        Database = database;
        Username = username;
        Password = password;
    }

    public NpgsqlConnection GetConnection()
    {
        try
        {
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