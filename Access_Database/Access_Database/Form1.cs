using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Access_Database
{
    public partial class Form1 : Form
    {
        //OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\NIKITA\access.accdb");
        OleDbConnection con;
        OleDbCommand cmd;
        OleDbDataAdapter da;
        DataSet ds;
        public Form1()
        {
            InitializeComponent();
        }
        void GetStudent()
        {
            con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\NIKITA\access.accdb");
            da = new OleDbDataAdapter("select * from student ", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "student");
            dataGridView1.DataSource = ds.Tables["student"];
            con.Close();
        }
        void Clear()
        {
            textBoxId.Text = "";
            textBoxName.Text = "";
            textBoxAddress.Text = "";
        }
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "select * from student where ID="+textBoxId.Text+"";
            OleDbDataReader dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                MessageBox.Show("Record already exists");
                con.Close();
            }
            else
            {
                con.Close();
                string command = "insert into student (ID,SName,Address) values(@ID,@SName,@Address)";
                cmd = new OleDbCommand(command, con);
                cmd.Parameters.AddWithValue("@ID", textBoxId.Text);
                cmd.Parameters.AddWithValue("@SName", textBoxName.Text);
                cmd.Parameters.AddWithValue("@Address", textBoxAddress.Text);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                GetStudent();
                Clear();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "MS Access DataBase";
            GetStudent();
        }

        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            string command = "update student set SName=@SName,Address=@Address where ID=@ID";
            cmd = new OleDbCommand(command, con);
            cmd.Parameters.AddWithValue("@SName", textBoxName.Text);
            cmd.Parameters.AddWithValue("@Address", textBoxAddress.Text);
            cmd.Parameters.AddWithValue("@ID", int.Parse(textBoxId.Text));
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            GetStudent();
            Clear();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            string command = "delete from student where ID="+textBoxId.Text+"";
            cmd = new OleDbCommand(command, con);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
            GetStudent();
            Clear();    
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                //gets a collection that contains all the rows
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                //populate the textbox from specific value of the coordinates of column and row.
                textBoxId.Text = row.Cells[0].Value.ToString();
                textBoxName.Text = row.Cells[1].Value.ToString();
                textBoxAddress.Text = row.Cells[2].Value.ToString();

            }
        }
    }
}
