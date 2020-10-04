using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using UserMaintenance.Entities;
using System.IO;

namespace UserMaintenance
{
    public partial class Form1 : Form
    {
        BindingList<User> users = new BindingList<User>();

        public Form1()
        {
            InitializeComponent();

            // Form Design
            label1.Text = Resource.FullName;
            button1.Text = Resource.Add;
            button2.Text = Resource.WriteToFile;
            button3.Text = Resource.Delete;

            // listBox1
            listBox1.DataSource = users;
            listBox1.ValueMember = "ID";
            listBox1.DisplayMember = "FullName";

            // Event Handlers
            button1.Click += addUser;
            button2.Click += writeToFile;
            button3.Click += deleteSelectedFromList;
        }

        private void addUser(object sender, EventArgs e)
        {
            var u = new User()
            {
                FullName = textBox1.Text
            };
            users.Add(u);
            textBox1.Clear();
        }

        private void writeToFile(object sender, EventArgs e) 
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog.FileName = Resource.DefaultFileName;
            saveFileDialog.DefaultExt = "txt";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(saveFileDialog.OpenFile());
                foreach (User u in users)
                {
                    writer.WriteLine(u.FullName);
                }
                writer.Close();
            }
        }

        private void deleteSelectedFromList(object sender, EventArgs e)
        {
            users.Remove((User)listBox1.SelectedItem);
        }
    }
}
