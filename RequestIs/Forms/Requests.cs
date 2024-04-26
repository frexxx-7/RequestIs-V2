using Guna.UI2.WinForms;
using MySql.Data.MySqlClient;
using RequestIs.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RequestIs.Forms
{
    public partial class Requests : Form
    {
        private Guna2DataGridView selectedDataGrid;
        private int selectedTab;
        private bool isEdit = false;
        public Requests()
        {
            InitializeComponent();
        }
        private void loadInfoCategory()
        {
            DB db = new DB();

            CategoryDataGrid.Rows.Clear();

            string query = $"select * from category ";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(query, db.getConnection()))
            {
                MySqlDataReader reader = mySqlCommand.ExecuteReader();

                List<string[]> dataDB = new List<string[]>();
                while (reader.Read())
                {

                    dataDB.Add(new string[reader.FieldCount]);

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dataDB[dataDB.Count - 1][i] = reader[i].ToString();
                    }
                }
                reader.Close();
                foreach (string[] s in dataDB)
                    CategoryDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void loadInfoStatusRequest()
        {
            DB db = new DB();

            StatusRequestDataGridView.Rows.Clear();

            string query = $"select * from statusrequest ";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(query, db.getConnection()))
            {
                MySqlDataReader reader = mySqlCommand.ExecuteReader();

                List<string[]> dataDB = new List<string[]>();
                while (reader.Read())
                {

                    dataDB.Add(new string[reader.FieldCount]);

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dataDB[dataDB.Count - 1][i] = reader[i].ToString();
                    }
                }
                reader.Close();
                foreach (string[] s in dataDB)
                    StatusRequestDataGridView.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void loadInfoUsers()
        {
            DB db = new DB();
            string queryInfo = $"SELECT id, concat(users.surname, ' ', users.name, ' ', users.patronymic) FROM users";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                UserComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoCategoryComboBox()
        {
            DB db = new DB();
            string queryInfo = $"SELECT id, name FROM category";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                CategoryComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }

        private void loadInfoRequests()
        {
            DB db = new DB();

            RequestsDataGrid.Rows.Clear();

            string query = $"select requests.id, requests.header, requests.content, concat(users.surname, ' ', users.name, ' ',users.patronymic) as FIOUser, category.name, dateRequest from requests " +
                $"join users on users.id = requests.idUser " +
                $"join category on category.id = requests.idCategory";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(query, db.getConnection()))
            {
                MySqlDataReader reader = mySqlCommand.ExecuteReader();

                List<string[]> dataDB = new List<string[]>();
                while (reader.Read())
                {

                    dataDB.Add(new string[reader.FieldCount]);

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dataDB[dataDB.Count - 1][i] = reader[i].ToString();
                    }
                }
                reader.Close();
                foreach (string[] s in dataDB)
                    RequestsDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void Requests_Load(object sender, EventArgs e)
        {
            selectedDataGrid = RequestsDataGrid;
            loadInfoRequests();
            loadInfoUsers();
            loadInfoCategoryComboBox();
            selectedTab = 0;
        }

        private void addRequestInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into requests (header, content, idUser, idCategory, dateRequest) values(@header, @content, @idUser, @idCategory, @dateRequest)", db.getConnection());
            command.Parameters.AddWithValue("@header", HeaderTextBox.Text);
            command.Parameters.AddWithValue("@content", ContentTextBox.Text);
            command.Parameters.AddWithValue("@idUser", (UserComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idCategory", (CategoryComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@dateRequest", dateRequestTimePicker.Value.ToString("yyyy.MM.dd"));
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Заявка добавлена");
                loadInfoRequests();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateRequestInDB(string idRequest)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update requests set header=@header, content=@content, idUser=@idUser, idCategory=@idCategory, dateRequest=@dateRequest where id = {idRequest}", db.getConnection());
            command.Parameters.AddWithValue("@header", HeaderTextBox.Text);
            command.Parameters.AddWithValue("@content", ContentTextBox.Text);
            command.Parameters.AddWithValue("@idUser", (UserComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idCategory", (CategoryComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@dateRequest", dateRequestTimePicker.Value.ToString("yyyy.MM.dd"));

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Заявка изменена");
                loadInfoRequests();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addCategoryInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into category (name) values(@name)", db.getConnection());
            command.Parameters.AddWithValue("@name", NameCategoryTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Категория добавлена");
                loadInfoCategory();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateCategoryInDB(string idCategory)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update category set name=@name where id = {idCategory}", db.getConnection());
            command.Parameters.AddWithValue("@name", NameCategoryTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Категория изменена");
                loadInfoCategory();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addStatusRequestInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into statusrequest (name) values(@name)", db.getConnection());
            command.Parameters.AddWithValue("@name", StatusRequestTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Статус обращения добавлен");
                loadInfoStatusRequest();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateStatusRequestInDB(string idStatusRequest)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update statusrequest set name=@name where id = {idStatusRequest}", db.getConnection());
            command.Parameters.AddWithValue("@name", StatusRequestTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Статус обращения изменен");
                loadInfoStatusRequest();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void AddButton_Click(object sender, EventArgs e)
        {
            if (selectedTab == 0)
            {
                if (!isEdit)
                {
                    addRequestInDB();
                }
                else
                {
                    updateRequestInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                }
            }
            else
            {
                if (selectedTab == 1)
                {
                    if (!isEdit)
                    {
                        addCategoryInDB();
                    }
                    else
                    {
                        updateCategoryInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    }
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        if (!isEdit)
                        {
                            addStatusRequestInDB();
                        }
                        else
                        {
                            updateStatusRequestInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                        }
                    }
                }
            }
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void UsersButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Users().Show();
        }

        private void MainButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Main().Show();
        }

        private void guna2TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedTab = guna2TabControl1.SelectedIndex;
            if (guna2TabControl1.SelectedIndex == 0)
            {
                selectedDataGrid = RequestsDataGrid;
                loadInfoRequests();
                loadInfoUsers();
                loadInfoCategoryComboBox();
            }
            else
            {
                if (guna2TabControl1.SelectedIndex == 1)
                {
                    selectedDataGrid = CategoryDataGrid;
                    loadInfoCategory();
                }
                else
                {
                    if (guna2TabControl1.SelectedIndex == 2)
                    {
                        selectedDataGrid = StatusRequestDataGridView;
                        loadInfoStatusRequest();
                    }
                }
            }
        }
    }
}
