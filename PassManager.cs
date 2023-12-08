using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using Konscious.Security.Cryptography;

namespace ShiftReportApp1
{
    public class PassManager
    {
        private const string ConfigFilePath = "sourseS.txt";
        private const int Iterations = 10; // Количество итераций

        public bool VerifyPassword(string inputPassword)
        {
            string savedHash = LoadHashFromConfig();
            string inputHash = GetPasswordHash(inputPassword);

            return savedHash == inputHash;
        }

        public void SetPassword(string newPassword)
        {
            string hash = GetPasswordHash(newPassword);
            SaveHashToConfig(hash);
        }

        private string GetPasswordHash(string password)
        {
            using (var hasher = new Argon2id(Encoding.UTF8.GetBytes(password)))
            {
                // Настраиваем параметры Argon2
                hasher.DegreeOfParallelism = 8; // По умолчанию 8
                hasher.MemorySize = 49152; // (в килобайтах)
                hasher.Iterations = Iterations;

                byte[] hashedBytes = hasher.GetBytes(32); // 32 байта для хранения хэша
                return BitConverter.ToString(hashedBytes).Replace("-", "");
            }
        }

        private void SaveHashToConfig(string hash)
        {
            try
            {
                File.WriteAllText(ConfigFilePath, hash);
            }
            catch (IOException ex)
            {
                // Обработка ошибки записи в файл
                Console.WriteLine($"Ошибка записи в файл: {ex.Message}");
            }
        }

        private string LoadHashFromConfig()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    return File.ReadAllText(ConfigFilePath);
                }
            }
            catch (IOException ex)
            {
                // Обработка ошибки чтения файла
                Console.WriteLine($"Ошибка чтения файла: {ex.Message}");
            }
            return null;
        }

        public void SetInitialPassword(string initialPassword)
        {
            // Вы можете установить первичный пароль только, если его еще нет
            if (LoadHashFromConfig() == null)
            {
                SetPassword(initialPassword);
            }
        }
    }
}