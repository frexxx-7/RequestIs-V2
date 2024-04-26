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
    public partial class Employee : Form
    {
        private Guna2DataGridView selectedDataGrid;
        private int selectedTab;
        private bool isEdit = false;
        public Employee()
        {
            InitializeComponent();
            EmployeeDataGrid.ClearSelection();
        }
        private void loadInfoPosition()
        {
            DB db = new DB();

            PositionsDataGridView.Rows.Clear();

            string query = $"select * from positions ";

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
                    PositionsDataGridView.Rows.Add(s);
            }
            db.closeConnection();
        }
        private void loadInfoEmployee()
        {
            DB db = new DB();

            EmployeeDataGrid.Rows.Clear();

            string query = $"select employee.id, employee.surname, employee.name, employee.patronymic, employee.numberPhone from employee";
                

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
                    EmployeeDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void UsersButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Users().Show();
        }

        private void Request_Click(object sender, EventArgs e)
        {
            this.Close();
            new Requests().Show();
        }

        private void Employee_Load(object sender, EventArgs e)
        {
            selectedDataGrid = EmployeeDataGrid;
            loadInfoEmployee();
            selectedTab = 0;
        }

        private void MainButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Main().Show();
        }
        private void addEmployeeInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into employee (surname, name, patronymic, numberPhone) values(@surname, @name, @patronymic, @numberPhone)", db.getConnection());
            command.Parameters.AddWithValue("@surname", SurnameTextBox.Text);
            command.Parameters.AddWithValue("@name", NameTextBox.Text);
            command.Parameters.AddWithValue("@patronymic", PatronymicTextBox.Text);
            command.Parameters.AddWithValue("@numberPhone", NumberPhoneTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Сотрудник добавлен");
                loadInfoEmployee();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateEmployeeInDB(string idEmployee)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update employee set surname=@surname, name=@name, patronymic=@patronymic, numberPhone=@numberPhone where id = {idEmployee}", db.getConnection());
            command.Parameters.AddWithValue("@surname", SurnameTextBox.Text);
            command.Parameters.AddWithValue("@name", NameTextBox.Text);
            command.Parameters.AddWithValue("@patronymic", PatronymicTextBox.Text);
            command.Parameters.AddWithValue("@numberPhone", NumberPhoneTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Сотрудник изменен");
                loadInfoEmployee();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addPositionInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into positions (name) values(@name)", db.getConnection());
            command.Parameters.AddWithValue("@name", NamePosTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Должность добавлена");
                loadInfoPosition();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updatePositionInDB(string idRegion)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update positions set name=@name where id = {idRegion}", db.getConnection());
            command.Parameters.AddWithValue("@name", NamePosTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Должность измененв");
                loadInfoPosition();

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
                    addEmployeeInDB();
                }
                else
                {
                    updateEmployeeInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                }
            }
            else
            {
                if (selectedTab == 1)
                {
                    if (!isEdit)
                    {
                        addPositionInDB();
                    }
                    else
                    {
                        updatePositionInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    }
                }
            }
            AddButton.Text = "Добавить";
            isEdit = false;
            if (selectedTab == 0)
            {
                SurnameTextBox.Text = "";
                NameTextBox.Text = "";
                PatronymicTextBox.Text = "";
                NumberPhoneTextBox.Text = "";
            }
            else
            {
                if (selectedTab == 1)
                {
                    NamePosTextBox.Text = "";
                }
            }
        }

        private void guna2TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedTab = guna2TabControl1.SelectedIndex;

            if (guna2TabControl1.SelectedIndex == 0)
            {
                selectedDataGrid = EmployeeDataGrid;
                loadInfoEmployee();
                ReportButton.Visible = true;
            }
            else
            {
                ReportButton.Visible = false;
                if (guna2TabControl1.SelectedIndex == 1)
                {
                    selectedDataGrid = PositionsDataGridView;
                    loadInfoPosition();
                }

            }
            selectedDataGrid.ClearSelection();
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void loadInfoOneEmployee(string idEmployee)
        {
            DB db = new DB();
            string queryInfo = $"select employee.id, employee.surname, employee.name, employee.patronymic, employee.numberPhone from employee " +
            $"where employee.id = {idEmployee}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                SurnameTextBox.Text = reader["surname"].ToString();
                NameTextBox.Text = reader["name"].ToString();
                PatronymicTextBox.Text = reader["patronymic"].ToString();
                NumberPhoneTextBox.Text = reader["numberPhone"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOnePostion(string idRegion)
        {
            DB db = new DB();
            string queryInfo = $"select * from positions " +
            $"where id = {idRegion}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                NamePosTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            isEdit = true;
            AddButton.Text = "Сохранить";
            if (selectedTab == 0)
            {
                loadInfoOneEmployee(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
            }
            else
            {
                if (selectedTab == 1)
                {
                    loadInfoOnePostion(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
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
                deleteRecordInBd("employee", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                loadInfoEmployee();
            }
            else
            {
                if (selectedTab == 1)
                {
                    deleteRecordInBd("positions", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    loadInfoPosition();
                }
            }
        }

        private void ReportButton_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            for (int j = 0; j < EmployeeDataGrid.Columns.Count; j++)
            {
                if (EmployeeDataGrid.Columns[j].Visible)
                {
                    worksheet.Cells[1, j] = EmployeeDataGrid.Columns[j].HeaderText;
                }
            }
            for (int i = 0; i < EmployeeDataGrid.Rows.Count; i++)
            {
                for (int j = 0; j < EmployeeDataGrid.Columns.Count; j++)
                {
                    if (EmployeeDataGrid.Columns[j].Visible)
                    {
                        worksheet.Cells[i + 2, j] = EmployeeDataGrid.Rows[i].Cells[j].Value;
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
    }
}
