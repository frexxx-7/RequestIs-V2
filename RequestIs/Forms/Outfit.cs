using Microsoft.Office.Interop.Word;
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
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace RequestIs.Forms
{
    public partial class Outfit : Form
    {
        public Outfit()
        {
            InitializeComponent();
        }
        private void loadInfoRequestsComboBox()
        {
            RequestComboBox.Items.Clear();

            DB db = new DB();
            string queryInfo = $"SELECT requests.id, requests.header FROM history " +
                $"inner join requests on requests.id = history.idRequest";
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
        private void Outfit_Load(object sender, EventArgs e)
        {
            loadInfoEmployeeComboBox();
            loadInfoRequestsComboBox();
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            DB db = new DB();
            MySqlCommand command = new MySqlCommand($"INSERT into outfit (idEmployee, idRequest) values(@idEmployee, @idRequest)", db.getConnection());
            command.Parameters.AddWithValue("@idEmployee", (EmployeeComboBox.SelectedItem as ComboBoxItem).Value);
            command.Parameters.AddWithValue("@idRequest", (RequestComboBox.SelectedItem as ComboBoxItem).Value);
            db.openConnection();

            try
            {
                command.ExecuteNonQuery();
                MessageBox.Show("Наряд добавлен");
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private string stateReq;
        private string content;
        private string inicUser;
        private string inicEmployee;
        private string reqheader;
        private string reqdateRequest;
        private void loadInfoVariable()
        {
            DB db = new DB();
            string queryInfo = $"select statusrequest.name, requests.content, concat(employee.surname, ' ', LEFT(employee.name,1), '.', LEFT(employee.patronymic,1)), concat(users.surname, ' ', LEFT(users.name,1), '.', LEFT(users.patronymic,1)), requests.header, requests.dateRequest from history " +
                $"inner join requests on history.idRequest = requests.id " +
                $"inner join users on requests.idUser = users.id " +
                $"inner join employee on history.idEmployee = employee.id " +
                $"inner join statusrequest on history.idStatusRequest = statusrequest.id " +
            $"where requests.id = {(RequestComboBox.SelectedItem as ComboBoxItem).Value}";
            MySqlCommand mySqlCommand = new MySqlCommand(queryInfo, db.getConnection());

            db.openConnection();

            MySqlDataReader reader = mySqlCommand.ExecuteReader();
            while (reader.Read())
            {
                stateReq = reader[0].ToString();
                content = reader[1].ToString();
                inicEmployee = reader[2].ToString();
                inicUser = reader[3].ToString();
                reqheader = reader[4].ToString();
                reqdateRequest = reader[5].ToString();
            }
            reader.Close();

            db.closeConnection();
        }
        private (string fullName, string date) LoadDirectorInfo()
        {
            string directorInfo = "";
            string date = DateTime.Now.ToString("dd.MM.yyyy");
            string query = "SELECT surname, name, patronymic FROM employee WHERE idPosition = (SELECT id FROM positions WHERE name = 'Директор') LIMIT 1";

            using (DB db = new DB())
            {
                db.openConnection();
                using (MySqlCommand command = new MySqlCommand(query, db.getConnection()))
                {
                    using (MySqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string surname = reader["surname"].ToString();
                            string name = reader["name"].ToString();
                            string patronymic = reader["patronymic"].ToString();
                            directorInfo = $"{surname} {name} {patronymic}";
                        }
                    }
                }
                db.closeConnection();
            }

            return (directorInfo, date);
        }
        private void ReportButton_Click(object sender, EventArgs e)
        {
            (string directorName, string currentDate) = LoadDirectorInfo();
            loadInfoVariable();
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Document sourceDoc = wordApp.Documents.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Шаблон наряд.docx"));
            sourceDoc.Content.Copy();

            Document targetDoc = wordApp.Documents.Add();
            targetDoc.Content.Paste();
            ReplaceBookmarkText(targetDoc, "ДатаСегодня", DateTime.Now.ToString("dd.MM.yyyy"));
            ReplaceBookmarkText(targetDoc, "НаименованиеЗаявки", reqheader);
            ReplaceBookmarkText(targetDoc, "Текст", content);
            ReplaceBookmarkText(targetDoc, "ИницПользователя", inicUser);
            ReplaceBookmarkText(targetDoc, "ИницСотрудника", inicEmployee);
            ReplaceBookmarkText(targetDoc, "ДатаОбращения", reqdateRequest);
            ReplaceBookmarkText(targetDoc, "ВыдалФИО", Main.fio);
            ReplaceBookmarkText(targetDoc, "ФИОДиректор", directorName);
            ReplaceBookmarkText(targetDoc, "ДатаСегодня2", currentDate);
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

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            new Requests().Show();
            this.Close();
        }
    }
}
