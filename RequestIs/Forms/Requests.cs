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
using Excel = Microsoft.Office.Interop.Excel;

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
            UserComboBox.Items.Clear();

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
            CategoryComboBox.Items.Clear();

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

            AddButton.Text = "Добавить";
            isEdit= false;
            if (selectedTab == 0)
            {
                HeaderTextBox.Text = "";
                ContentTextBox.Text = "";
                UserComboBox.SelectedIndex = -1;
                CategoryComboBox.SelectedIndex = -1;
                dateRequestTimePicker.Value = DateTime.Now;
            }
            else
            {
                if (selectedTab == 1)
                {
                    NameCategoryTextBox.Text = "";
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        StatusRequestTextBox.Text = "";
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
                ReportButton.Visible = true;
            }
            else
            {
                ReportButton.Visible = false;
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
            selectedDataGrid.ClearSelection();
        }
        private void loadInfoOneRequest(string idRequest)
        {
            DB db = new DB();
            string queryInfo = $"select requests.id, requests.idUser, requests.idCategory, requests.header, requests.content, concat(users.surname, ' ', users.name, ' ',users.patronymic) as FIOUser, category.name, dateRequest from requests " +
                $"join users on users.id = requests.idUser " +
                $"join category on category.id = requests.idCategory " +
                $"where requests.id = {idRequest}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                HeaderTextBox.Text = reader["header"].ToString();
                ContentTextBox.Text = reader["content"].ToString();
                for (int i = 0; i < UserComboBox.Items.Count; i++)
                {
                    if (reader["idUser"].ToString() != "")
                    {
                        if (Convert.ToInt32((UserComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idUser"]))
                        {
                            UserComboBox.SelectedIndex = i;
                        }
                    }
                }
                for (int i = 0; i < CategoryComboBox.Items.Count; i++)
                {
                    if (reader["idCategory"].ToString() != "")
                    {
                        if (Convert.ToInt32((CategoryComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idCategory"]))
                        {
                            CategoryComboBox.SelectedIndex = i;
                        }
                    }
                }
                dateRequestTimePicker.Value = Convert.ToDateTime(reader["dateRequest"].ToString());
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOneCategory(string idCategory)
        {
            DB db = new DB();
            string queryInfo = $"select * from category " +
                $"where category.id = {idCategory}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                NameCategoryTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOneStatus(string idStatus)
        {
            DB db = new DB();
            string queryInfo = $"select * from statusrequest " +
                $"where statusrequest.id = {idStatus}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                StatusRequestTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void EditButton_Click(object sender, EventArgs e)
        {
            isEdit = true;
            AddButton.Text = "Сохранить";
            if(selectedTab == 0)
            {
                loadInfoOneRequest(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
            }
            else
            {
                if (selectedTab == 1)
                {
                    loadInfoOneCategory(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        loadInfoOneStatus(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    }
                }
            }
        }
        private void deleteRecordInBd(string tableName, string id)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"delete from {tableName} where id = {id}", db.getConnection());
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Запись удалена");

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (selectedTab == 0)
            {
                deleteRecordInBd("requests", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                loadInfoRequests();
            }
            else
            {
                if (selectedTab == 1)
                {
                    deleteRecordInBd("category", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    loadInfoCategory();
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        deleteRecordInBd("statusrequest", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                        loadInfoStatusRequest();
                    }
                }
            }
        }

        private void ReportButton_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            for (int j = 0; j < RequestsDataGrid.Columns.Count; j++)
            {
                if (RequestsDataGrid.Columns[j].Visible)
                {
                    worksheet.Cells[1, j] = RequestsDataGrid.Columns[j].HeaderText;
                }
            }
            for (int i = 0; i < RequestsDataGrid.Rows.Count; i++)
            {
                for (int j = 0; j < RequestsDataGrid.Columns.Count; j++)
                {
                    if (RequestsDataGrid.Columns[j].Visible)
                    {
                        worksheet.Cells[i + 2, j] = RequestsDataGrid.Rows[i].Cells[j].Value;
                    }
                }
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File|*.xlsx";
            saveFileDialog1.Title = "Сохранить Excel файл";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                workbook.SaveAs(saveFileDialog1.FileName);
            }
            workbook.Close();
            excelApp.Quit();
        }

        private void EmployeeButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Employee().Show();
        }
    }
}
