using Guna.UI2.WinForms;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using RequestIs.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

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
        private void loadInfoHistoryRequests()
        {
            DB db = new DB();

            HistoryRequestDataGridView.Rows.Clear();

            string query = $"select history.id, requests.header, statusrequest.name, history.dateEdit, concat(employee.surname, ' ', employee.name, ' ', employee.patronymic) from history " +
                $"inner join requests on history.idRequest = requests.id " +
                $"inner join statusrequest on history.idStatusRequest = statusrequest.id " +
                $"inner join employee on history.idEmployee = employee.id ";

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
                    HistoryRequestDataGridView.Rows.Add(s);
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
        private void loadInfoRequestsComboBox()
        {
            RequestComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT id, header FROM requests";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                RequestComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoStateComboBox()
        {
            StateRequestComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT id, name FROM statusrequest";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                StateRequestComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoEmployeeComboBox()
        {
            EmployeeComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT id, name FROM employee";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                EmployeeComboBox.Items.Add(item);
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
            loadInfoHistoryRequests();
            selectedTab = 0;
        }
        private void addRequestInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into requests (header, content, idUser, idCategory, dateRequest) values(@header, @content, @idUser, @idCategory, @dateRequest); SELECT LAST_INSERT_ID();", db.getConnection());
            command.Parameters.AddWithValue("@header", HeaderTextBox.Text);
            command.Parameters.AddWithValue("@content", ContentTextBox.Text);
            command.Parameters.AddWithValue("@idUser", (UserComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idCategory", (CategoryComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@dateRequest", dateRequestTimePicker.Value.ToString("dd.MM.yyyy"));
            db.openConnection();

            try
            {
                int newRequestId = Convert.ToInt32(command.ExecuteScalar());
                MessageBox.Show("Заявка добавлена");
                loadInfoRequests();
                addHisotryInDBForRequest(newRequestId);
            }
            catch(Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }

        private void addHisotryInDBForRequest(int requestId)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into history (idRequest, idStatusRequest, dateEdit, idEmployee) values(@idRequest, @idStatusRequest, @dateEdit, @idEmployee)", db.getConnection());
            command.Parameters.AddWithValue("@idRequest", requestId);
            command.Parameters.AddWithValue("@idStatusRequest", 1);
            command.Parameters.AddWithValue("@dateEdit", DateTime.Now.ToString("dd.MM.yyyy"));
            command.Parameters.AddWithValue("@idEmployee", 1);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("История добавлена");
                loadInfoHistoryRequests();
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            command.Parameters.AddWithValue("@dateRequest", dateRequestTimePicker.Value.ToString("dd.MM.yyyy"));

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
        private void addHisotryInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into history (idRequest, idStatusRequest, dateEdit, idEmployee) values(@idRequest, @idStatusRequest, @dateEdit, @idEmployee)", db.getConnection());
            command.Parameters.AddWithValue("@idRequest", (RequestComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idStatusRequest", (StateRequestComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@dateEdit", DateEditDateTimePicker.Value.ToString("dd.MM.yyyy"));
            command.Parameters.AddWithValue("@idEmployee", (EmployeeComboBox.SelectedItem as ComboBoxItem).Value);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("История добавлена");
                loadInfoHistoryRequests();
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
        private void updateHistoryInDB(string idStatusRequest)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update history set idRequest=@idRequest, idStatusRequest=@idStatusRequest, dateEdit=@dateEdit, idEmployee=@idEmployee where id = {idStatusRequest}", db.getConnection());
            command.Parameters.AddWithValue("@idRequest", (RequestComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idStatusRequest", (StateRequestComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@dateEdit", DateEditDateTimePicker.Value.ToString("yyyy.MM.dd"));
            command.Parameters.AddWithValue("@idEmployee", (EmployeeComboBox.SelectedItem as ComboBoxItem).Value);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("История изменена");
                loadInfoHistoryRequests();

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
                    else
                    {
                        if (selectedTab == 3)
                        {
                            if (!isEdit)
                            {
                                addHisotryInDB();
                            }
                            else
                            {
                                updateHistoryInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                            }
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
                    else
                    {
                        if (selectedTab == 2)
                        {
                            DateEditDateTimePicker.Value = DateTime.Now;
                            RequestComboBox.SelectedIndex = -1;
                            StateRequestComboBox.SelectedIndex = -1;
                            EmployeeComboBox.SelectedIndex = -1;
                        }
                    }
                }
            }
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
           System.Windows.Forms.Application.Exit();
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
                ReportWordButton.Visible = true;
            }
            else
            {
                ReportButton.Visible = false;
                ReportWordButton.Visible = false;
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
                    else
                    {
                        if (guna2TabControl1.SelectedIndex == 3)
                        {
                            selectedDataGrid = HistoryRequestDataGridView;
                            loadInfoHistoryRequests();
                            loadInfoEmployeeComboBox();
                            loadInfoRequestsComboBox();
                            loadInfoStateComboBox();
                        }
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
        private void loadInfoOneHistory(string idStatus)
        {
            DB db = new DB();
            string queryInfo = $"select * from history " +
                $"where history.id = {idStatus}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                for (int i = 0; i < RequestComboBox.Items.Count; i++)
                {
                    if (reader["idRequest"].ToString() != "")
                    {
                        if (Convert.ToInt32((RequestComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idRequest"]))
                        {
                            RequestComboBox.SelectedIndex = i;
                        }
                    }
                }
                for (int i = 0; i < StateRequestComboBox.Items.Count; i++)
                {
                    if (reader["idStatusRequest"].ToString() != "")
                    {
                        if (Convert.ToInt32((StateRequestComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idStatusRequest"]))
                        {
                            StateRequestComboBox.SelectedIndex = i;
                        }
                    }
                }
                for (int i = 0; i < EmployeeComboBox.Items.Count; i++)
                {
                    if (reader["idEmployee"].ToString() != "")
                    {
                        if (Convert.ToInt32((EmployeeComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idEmployee"]))
                        {
                            EmployeeComboBox.SelectedIndex = i;
                        }
                    }
                }
                DateEditDateTimePicker.Value = Convert.ToDateTime(reader["dateEdit"].ToString());
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
                    else
                    {
                        if (selectedTab == 3)
                        {
                            loadInfoOneHistory(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                        }
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
                    else
                    {
                        if (selectedTab == 3)
                        {
                            deleteRecordInBd("history", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                            loadInfoHistoryRequests();
                        }
                    }
                }
            }
        }

        private void ReportButton_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            string fileName = "Отчет о заявках";
            int visibleColumnCount = 0;

            // Calculate the number of visible columns
            for (int j = 0; j < RequestsDataGrid.Columns.Count; j++)
            {
                if (RequestsDataGrid.Columns[j].Visible)
                {
                    visibleColumnCount++;
                }
            }

            Excel.Range titleRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, visibleColumnCount]];
            titleRange.Merge();
            titleRange.Value = fileName;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            int headerRow = 2;
            int dataStartRow = 3;

            int columnIndex = 1;
            for (int j = 0; j < RequestsDataGrid.Columns.Count; j++)
            {
                if (RequestsDataGrid.Columns[j].Visible)
                {
                    worksheet.Cells[headerRow, columnIndex] = RequestsDataGrid.Columns[j].HeaderText;
                    columnIndex++;
                }
            }

            for (int i = 0; i < RequestsDataGrid.Rows.Count; i++)
            {
                columnIndex = 1;
                for (int j = 0; j < RequestsDataGrid.Columns.Count; j++)
                {
                    if (RequestsDataGrid.Columns[j].Visible)
                    {
                        worksheet.Cells[i + dataStartRow, columnIndex] = RequestsDataGrid.Rows[i].Cells[j].Value;
                        columnIndex++;
                    }
                }
            }

            worksheet.Columns.AutoFit();

            Excel.Range usedRange = worksheet.UsedRange;
            usedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel File|*.xlsx";
            saveFileDialog1.Title = "Сохранить Excel файл";
            saveFileDialog1.FileName = fileName;
            saveFileDialog1.ShowDialog();

            if (saveFileDialog1.FileName != "")
            {
                workbook.SaveAs(saveFileDialog1.FileName);

                workbook.Close(false);
                excelApp.Quit();

                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            worksheet = null;
            workbook = null;
            excelApp = null;
            GC.Collect();
        }

        private void EmployeeButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Employee().Show();
        }

        private void SeacrhTextBox_TextChanged(object sender, EventArgs e)
        {
            DB db = new DB();

            selectedDataGrid.Rows.Clear();

            string searchString = $"select requests.id, requests.header, requests.content, concat(users.surname, ' ', users.name, ' ',users.patronymic) as FIOUser, category.name, dateRequest from requests " +
                $"join users on users.id = requests.idUser " +
                $"join category on category.id = requests.idCategory " +
                $"where concat (requests.header, requests.content, concat(users.surname, ' ', users.name, ' ',users.patronymic), category.name, dateRequest) like '%" + SeacrhTextBox.Text + "%'";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(searchString, db.getConnection()))
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
                    selectedDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            DB db = new DB();

            selectedDataGrid.Rows.Clear();
            string searchString = $"select * from category " +
                $"where concat (category.name) like '%" + guna2TextBox1.Text + "%'";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(searchString, db.getConnection()))
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
                    selectedDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {
            DB db = new DB();

            selectedDataGrid.Rows.Clear();

            string searchString = $"select * from statusrequest " +
                $"where concat (name) like '%" + guna2TextBox2.Text + "%'";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(searchString, db.getConnection()))
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
                    selectedDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void guna2TextBox3_TextChanged(object sender, EventArgs e)
        {
            DB db = new DB();

            selectedDataGrid.Rows.Clear();

            string searchString = $"select history.id, requests.header, statusrequest.name, history.dateEdit, concat(employee.surname, ' ', employee.name, ' ', employee.patronymic) from history " +
                $"inner join requests on history.idRequest = requests.id " +
                $"inner join statusrequest on history.idStatusRequest = statusrequest.id " +
                $"inner join employee on history.idEmployee = employee.id "+
            $"where concat (requests.header, statusrequest.name, history.dateEdit, concat(employee.surname, ' ', employee.name, ' ', employee.patronymic)) like '%" + guna2TextBox3.Text + "%'";

            db.openConnection();
            using (MySqlCommand mySqlCommand = new MySqlCommand(searchString, db.getConnection()))
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
                    selectedDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }
        private void ReplaceBookmarkText(Document doc, string bookmarkName, string text)
        {
            Bookmark bookmark = doc.Bookmarks[bookmarkName];
            if (bookmark != null)
            {
                Microsoft.Office.Interop.Word.Range range = bookmark.Range;
                range.Text = text;
            }
        }
        private string reqdateRequest;
        private string dateEdit;
        private string reqheader;
        private string content;
        private string fioUser;
        private string inicUser;
        private string inicEmployee;
        private void loadInfoVariable()
        {
            DB db = new DB();
            string queryInfo = $"select requests.dateRequest, history.dateEdit, requests.header, requests.content, concat(users.surname, ' ', users.name, ' ', users.patronymic), concat(employee.surname, ' ', LEFT(employee.name,1), '.', LEFT(employee.patronymic,1)), concat(users.surname, ' ', LEFT(users.name,1), '.', LEFT(users.patronymic,1)) from history " +
                $"inner join requests on history.idRequest = requests.id " +
                $"inner join users on requests.idUser = users.id " +
                $"inner join employee on history.idEmployee = employee.id " +
            $"where requests.id = {selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString()}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                reqdateRequest = reader[0].ToString();
                dateEdit = reader[1].ToString();
                reqheader = reader[2].ToString();
                content = reader[3].ToString();
                fioUser = reader[4].ToString();
                inicEmployee = reader[5].ToString();
                inicUser = reader[6].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void ReportWordButton_Click(object sender, EventArgs e)
        {
            loadInfoVariable();
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Document sourceDoc = wordApp.Documents.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Шаблон акт выполненных работ.docx"));
            sourceDoc.Content.Copy();

            Document targetDoc = wordApp.Documents.Add();
            targetDoc.Content.Paste();
            ReplaceBookmarkText(targetDoc, "ДатаОбращения", reqdateRequest);
            ReplaceBookmarkText(targetDoc, "ДатаОбращения2", reqdateRequest);
            ReplaceBookmarkText(targetDoc, "ДатаРедактирования", dateEdit);
            ReplaceBookmarkText(targetDoc, "ДатаСегодня", DateTime.Now.ToString("dd.MM.yyyy"));
            ReplaceBookmarkText(targetDoc, "НаименованиеЗаявки", reqheader);
            ReplaceBookmarkText(targetDoc, "Текст", content);
            ReplaceBookmarkText(targetDoc, "ФИОПользователя", fioUser);
            ReplaceBookmarkText(targetDoc, "ИницПользователя", inicUser);
            ReplaceBookmarkText(targetDoc, "ИницСотрудника", inicEmployee);

            sourceDoc.Close();

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = "Документ Word (*.docx)|*.docx",
                Title = "Сохранить скопированный документ в"
            };

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string targetPath = saveFileDialog1.FileName;
                targetDoc.SaveAs2(targetPath);
                targetDoc.Close();

                Document wordDocument = wordApp.Documents.Open(targetPath);
                wordApp.Visible = true;
            }
            else
            {
                targetDoc.Close(false);
                wordApp.Quit();
            }
        }
    }
}
