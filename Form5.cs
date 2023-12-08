using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class Form5 : Form
    {
        private PassManager passManager;

        public Form5()
        {
            InitializeComponent();
            passManager = new PassManager();

            // Установите первичный пароль при первом запуске
            passManager.SetInitialPassword("Manager");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void Form5_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Проверка правильности пароля
            if (passManager.VerifyPassword(textBox2.Text))
            {
                // Если пароль верный, создаем и открываем Form6
                Form6 form6 = new Form6();
                form6.Show();

                // Закрываем текущую форму
                this.Close();
            }
            else
            {
                // Если пароль неверный, вы можете показать сообщение об ошибке
                MessageBox.Show("Неверный пароль. Попробуйте снова.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
