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
        }

        private void guna2TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedTab = guna2TabControl1.SelectedIndex;
            if (guna2TabControl1.SelectedIndex == 0)
            {
                selectedDataGrid = UsersDataGrid;
                loadInfoLocalityComboBox();
                loadInfoUsers();
            }
            else
            {
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
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
