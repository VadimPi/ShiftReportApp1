using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Npgsql;

public class DBConnection
{
    private string Host { get; }
    private int Port { get; }
    private string Database { get; }
    private string Username { get; }
    private string Password { get; }

    public DBConnection(string host = "localhost", int port = 5432, string database = "reportdb", string username = "User", string password = "UserPassword")
    {
        Host = host;
        Port = port;
        Database = database;
        Username = username;
        Password = password;
    }

    public NpgsqlConnection GetConnection()
    {
        string connectionString = $"Host={Host};Port={Port};Database={Database};Username={Username};Password={Password}";
        return new NpgsqlConnection(connectionString);
    }
}
