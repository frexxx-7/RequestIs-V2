using MySql.Data.MySqlClient;
using RequestIs.Classes;
using RequestIs.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RequestIs
{
    public partial class Autorization : Form
    {
        public Autorization()
        {
            InitializeComponent();
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void RegisterButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            new Registration().Show();
        }

        private void AutorizationButton_Click(object sender, EventArgs e)
        {
            DB db = new DB();
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter();
            MySqlCommand command = new MySqlCommand("SELECT id, login, password, concat(employee.surname, ' ', LEFT(employee.name,1), '.', LEFT(employee.patronymic,1)) FROM employee WHERE login = @login AND password = @password", db.getConnection());

            command.Parameters.Add("@login", MySqlDbType.VarChar).Value = LoginTextBox.Text;
            command.Parameters.Add("@password", MySqlDbType.VarChar).Value = PasswordTextBox.Text;

            adapter.SelectCommand = command;
            adapter.Fill(table);
            if (table.Rows.Count > 0)
            {
                Main.idUser = table.Rows[0]["id"].ToString();
                Main.fio = table.Rows[0][3].ToString();
                Main main = new Main();
                this.Hide();
                main.Show();
                MessageBox.Show("Добро пожаловать");
            }
            else
            {
                MessageBox.Show("Неправильный логин или пароль");
            }
        }
    }
}
