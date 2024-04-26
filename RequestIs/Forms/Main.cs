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
    public partial class Main : Form
    {
        public static string idUser;
        public Main()
        {
            InitializeComponent();
        }

        private void guna2ControlBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void MainButton_Click(object sender, EventArgs e)
        {

        }

        private void Request_Click(object sender, EventArgs e)
        {
            this.Hide();
            new Requests().Show();
        }

        private void UsersButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            new Users().Show();
        }
    }
}
