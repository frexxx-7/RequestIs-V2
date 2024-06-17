using Guna.UI2.WinForms;
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
using System.Web.Routing;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace RequestIs.Forms
{
    public partial class Users : Form
    {
        private Guna2DataGridView selectedDataGrid;
        private int selectedTab;
        private bool isEdit = false;
        public Users()
        {
            InitializeComponent();
            UsersDataGrid.ClearSelection();
        }
        private void loadInfoRegions()
        {
            DB db = new DB();

            RegionDataGrid.Rows.Clear();

            string query = $"select region.id,  region.name from region ";

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
                    RegionDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }
        private void loadInfoArea()
        {
            DB db = new DB();

            AreaDataGrid.Rows.Clear();

            string query = $"select area.id,  region.name, area.name from area " +
                $"join region on region.id = area.idRegion";

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
                    AreaDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }
        private void loadInfoRegionComboBox()
        {
            RegionComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT id, name FROM region";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                RegionComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoLocality()
        {
            DB db = new DB();

            LocalityDataGridView.Rows.Clear();

            string query = $"select locality.id, concat(area.name, ' ', region.name) as address, locality.type, locality.name from locality " +
                $"join area on area.id = locality.idArea " +
                $"join region on region.id = area.idRegion";

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
                    LocalityDataGridView.Rows.Add(s);
            }
            db.closeConnection();
        }
        private void loadInfoAreaComboBox()
        {
            AreaComboBox.Items.Clear();
            DB db = new DB();
            string queryInfo = $"SELECT id, name FROM area";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                AreaComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoLocalityComboBox()
        {
            LocalityComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT locality.id, concat(area.name, ' ', region.name, ' ', locality.name, ' ') as address FROM locality " +
                $"join area on area.id = locality.idArea " +
                $"join region on region.id = area.idRegion";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Text = $" {reader[1]}";
                item.Value = reader[0];
                LocalityComboBox.Items.Add(item);
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoUsers()
        {
            DB db = new DB();

            UsersDataGrid.Rows.Clear();

            string query = $"select users.id, users.surname, users.name, users.patronymic, users.numberPhone, " +
                $"users.email, concat(area.name, ' ', region.name, ' ', locality.name, ' ') as address, users.street, users.house, users.apartment from users " +
                $"join locality on locality.id = users.idLocality " +
                $"join area on area.id = locality.idArea " +
                $"join region on region.id = area.idRegion";

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
                    UsersDataGrid.Rows.Add(s);
            }
            db.closeConnection();
        }

        private void Request_Click(object sender, EventArgs e)
        {
            this.Close();
            new Requests().Show();
        }

        private void MainButton_Click(object sender, EventArgs e)
        {
            this.Close();
            new Main().Show();
        }

        private void Users_Load(object sender, EventArgs e)
        {
            selectedDataGrid = UsersDataGrid;
            loadInfoUsers();
            loadInfoLocalityComboBox();
            selectedTab = 0;
        }
        private void addUserInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into users (surname, name, patronymic, numberPhone, email, idLocality, street, house, apartment) values(@surname, @name, @patronymic, @numberPhone, @email, @idLocality, @street, @house, @apartment)", db.getConnection());
            command.Parameters.AddWithValue("@surname", SurnameTextBox.Text);
            command.Parameters.AddWithValue("@name", NameTextBox.Text);
            command.Parameters.AddWithValue("@patronymic", PatronymicTextBox.Text);
            command.Parameters.AddWithValue("@numberPhone", NumberPhoneTextBox.Text);
            command.Parameters.AddWithValue("@email", EmailTextBox.Text);
            command.Parameters.AddWithValue("@idLocality", (LocalityComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@street", StreetTextBox.Text);
            command.Parameters.AddWithValue("@house", HouseTextBox.Text);
            command.Parameters.AddWithValue("@apartment", ApartmentTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь добавлен");
                loadInfoUsers();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateUserInDB(string idUser)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update users set surname=@surname, name=@name, patronymic=@patronymic, numberPhone=@numberPhone, email=@email, idLocality=@idLocality, street=@street, house=@house, apartment=@apartment where id = {idUser}", db.getConnection());
            command.Parameters.AddWithValue("@surname", SurnameTextBox.Text);
            command.Parameters.AddWithValue("@name", NameTextBox.Text);
            command.Parameters.AddWithValue("@patronymic", PatronymicTextBox.Text);
            command.Parameters.AddWithValue("@numberPhone", NumberPhoneTextBox.Text);
            command.Parameters.AddWithValue("@email", EmailTextBox.Text);
            command.Parameters.AddWithValue("@idLocality", (LocalityComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@street", StreetTextBox.Text);
            command.Parameters.AddWithValue("@house", HouseTextBox.Text);
            command.Parameters.AddWithValue("@apartment", ApartmentTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь изменен");
                loadInfoUsers();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addLocalityInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into locality (idArea, type, name) values(@idArea, @type, @name)", db.getConnection());
            command.Parameters.AddWithValue("@idArea", (AreaComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@type", TypeTextBox.Text);
            command.Parameters.AddWithValue("@name", NameLocTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Населенный пункт добавлен");
                loadInfoLocality();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateLocalityInDB(string idLocality)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update locality set idArea=@idArea, type=@type, name=@name where id = {idLocality}", db.getConnection());
            command.Parameters.AddWithValue("@idArea", (AreaComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@type", TypeTextBox.Text);
            command.Parameters.AddWithValue("@name", NameLocTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Населенный пункт изменен");
                loadInfoLocality();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addAreaInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into area (idRegion, name) values(@idRegion, @name)", db.getConnection());
            command.Parameters.AddWithValue("@idRegion", (RegionComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@name", NameAreaTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Область добавлена");
                loadInfoArea();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateAreaInDB(string idArea)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update area set idRegion=@idRegion, name=@name where id = {idArea}", db.getConnection());
            command.Parameters.AddWithValue("@idArea", (AreaComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@name", NameLocTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Область изменена");
                loadInfoArea();

            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void addRegionInDB()
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into region (name) values(@name)", db.getConnection());
            command.Parameters.AddWithValue("@name", NameRegionTextBox.Text);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Регион добавлен");
                loadInfoRegions();
            }
            catch
            {
                MessageBox.Show("Ошибка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            db.closeConnection();
        }
        private void updateRegionInDB(string idRegion)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"update region set name=@name where id = {idRegion}", db.getConnection());
            command.Parameters.AddWithValue("@name", NameRegionTextBox.Text);

            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Регион изменен");
                loadInfoArea();

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
                    addUserInDB();
                }
                else
                {
                    updateUserInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                }
            }
            else
            {
                if (selectedTab == 1)
                {
                    if (!isEdit)
                    {
                        addLocalityInDB();
                    }
                    else
                    {
                        updateLocalityInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    }
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        if (!isEdit)
                        {
                            addAreaInDB();
                        }
                        else
                        {
                            updateAreaInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                        }
                    }
                    else
                    {
                        if (selectedTab == 3)
                        {
                            if (!isEdit)
                            {
                                addRegionInDB();
                            }
                            else
                            {
                                updateRegionInDB(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                            }
                        }
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
                EmailTextBox.Text = "";
                LocalityComboBox.SelectedIndex = -1;
                StreetTextBox.Text = "";
                HouseTextBox.Text = "";
                ApartmentTextBox.Text = "";
            }
            else
            {
                if (selectedTab == 1)
                {
                    AreaComboBox.SelectedIndex = -1;
                    TypeTextBox.Text = "";
                    NameLocTextBox.Text= "";
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        RegionComboBox.SelectedIndex = -1;
                        NameAreaTextBox.Text = "";
                    }
                    else
                    {
                        if (selectedTab == 3)
                        {
                            NameRegionTextBox.Text = "";
                        }
                    }
                }
            }
        }

        private void guna2TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedTab = guna2TabControl1.SelectedIndex;
            
            if (guna2TabControl1.SelectedIndex == 0)
            {
                selectedDataGrid = UsersDataGrid;
                loadInfoLocalityComboBox();
                loadInfoUsers();
                ReportButton.Visible = true;
            }
            else
            {
                ReportButton.Visible = false;
                if (guna2TabControl1.SelectedIndex == 1)
                {
                    selectedDataGrid = LocalityDataGridView;
                    loadInfoAreaComboBox();
                    loadInfoLocality();
                }
                else
                {
                    if (guna2TabControl1.SelectedIndex == 2)
                    {
                        selectedDataGrid = AreaDataGrid;
                        loadInfoRegionComboBox();
                        loadInfoArea();
                    }
                    else
                    {
                        if (guna2TabControl1.SelectedIndex == 3)
                        {
                            selectedDataGrid = RegionDataGrid;
                            loadInfoRegions();
                        }
                    }
                }
            }

            selectedDataGrid.ClearSelection();
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void loadInfoOneUser(string idUser)
        {
            DB db = new DB();
            string queryInfo = $"select users.id, idLocality, users.surname, users.name, users.patronymic, users.numberPhone, " +
                $"users.email, concat(area.name, ' ', region.name, ' ', locality.name, ' ') as address, users.street, users.house, users.apartment from users " +
                $"join locality on locality.id = users.idLocality " +
                $"join area on area.id = locality.idArea " +
                $"join region on region.id = area.idRegion " +
            $"where users.id = {idUser}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                SurnameTextBox.Text = reader["surname"].ToString();
                NameTextBox.Text = reader["name"].ToString();
                PatronymicTextBox.Text = reader["patronymic"].ToString();
                NumberPhoneTextBox.Text = reader["numberPhone"].ToString();
                EmailTextBox.Text = reader["email"].ToString();

                for (int i = 0; i < LocalityComboBox.Items.Count; i++)
                {
                    if (reader["idLocality"].ToString() != "")
                    {
                        if (Convert.ToInt32((LocalityComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idLocality"]))
                        {
                            LocalityComboBox.SelectedIndex = i;
                        }
                    }
                }

                StreetTextBox.Text = reader["street"].ToString();
                HouseTextBox.Text = reader["house"].ToString();
                ApartmentTextBox.Text = reader["apartment"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOneLocality(string idLocality)
        {
            DB db = new DB();
            string queryInfo = $"select * from locality " +
            $"where locality.id = {idLocality}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                for (int i = 0; i < AreaComboBox.Items.Count; i++)
                {
                    if (reader["idArea"].ToString() != "")
                    {
                        if (Convert.ToInt32((AreaComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idArea"]))
                        {
                            AreaComboBox.SelectedIndex = i;
                        }
                    }
                }
                TypeTextBox.Text = reader["type"].ToString();
                NameLocTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOneArea(string idArea)
        {
            DB db = new DB();
            string queryInfo = $"select * from area " +
            $"where area.id = {idArea}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                for (int i = 0; i < RegionComboBox.Items.Count; i++)
                {
                    if (reader["idRegion"].ToString() != "")
                    {
                        if (Convert.ToInt32((RegionComboBox.Items[i] as ComboBoxItem).Value) == Convert.ToInt32(reader["idRegion"]))
                        {
                            RegionComboBox.SelectedIndex = i;
                        }
                    }
                }
                NameAreaTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private void loadInfoOneRegion(string idRegion)
        {
            DB db = new DB();
            string queryInfo = $"select * from region " +
            $"where id = {idRegion}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                NameRegionTextBox.Text = reader["name"].ToString();
            }
            reader.Close();

            db.closeConnection();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            isEdit = true;
            AddButton.Text = "Сохранить";
            if (selectedTab == 0)
            {
                loadInfoOneUser(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
            }
            else
            {
                if (selectedTab == 1)
                {
                    loadInfoOneLocality(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        loadInfoOneArea(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    }
                    else
                    {
                        if (selectedTab == 3)
                        {
                            loadInfoOneRegion(selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
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
                deleteRecordInBd("users", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                loadInfoUsers();
            }
            else
            {
                if (selectedTab == 1)
                {
                    deleteRecordInBd("locality", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                    loadInfoLocality();
                }
                else
                {
                    if (selectedTab == 2)
                    {
                        deleteRecordInBd("area", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                        loadInfoArea();
                    }
                    else
                    {
                        if (selectedTab == 3)
                        {
                            deleteRecordInBd("region", selectedDataGrid[0, selectedDataGrid.SelectedCells[0].RowIndex].Value.ToString());
                            loadInfoRegions();
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

            int columnIndex = 1; 
            for (int j = 0; j < UsersDataGrid.Columns.Count; j++)
            {
                if (UsersDataGrid.Columns[j].Visible)
                {
                    worksheet.Cells[1, columnIndex] = UsersDataGrid.Columns[j].HeaderText;
                    columnIndex++;
                }
            }

            for (int i = 0; i < UsersDataGrid.Rows.Count; i++)
            {
                columnIndex = 1; 
                for (int j = 0; j < UsersDataGrid.Columns.Count; j++)
                {
                    if (UsersDataGrid.Columns[j].Visible)
                    {
                        worksheet.Cells[i + 2, columnIndex] = UsersDataGrid.Rows[i].Cells[j].Value;
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
            saveFileDialog1.FileName = "Отчет о пользователях";
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
    }
}
