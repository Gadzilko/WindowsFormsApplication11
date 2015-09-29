using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form4 : Form
    {
        public int selectrow = -1;
        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1 = "SELECT * FROM [Клиенты]";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            if (ds.Tables["Результат"].Rows.Count > 0)
            {
                dataGridView1.Columns[0].Width = 80;
                dataGridView1.Columns[1].Width = 210;
            }

            this.Height = 310;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer3.Visible = false;
            button9.Visible = false;
            this.Height = 480;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            this.Height = 310;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1)
            {
                MessageBox.Show("Выделите в сетке строку для редактирования");
                return;
            }

            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            
            string KodKlienta = textBox4.Text.ToString();
            string FIO        = textBox5.Text.ToString();
            string Adres      = textBox6.Text.ToString();
            string Telephone  = textBox7.Text.ToString();

            string sq1 = " UPDATE Клиенты SET " +
                  " ФИО = '" + FIO + "' " + ", " +
                  " Адрес = '" + Adres + "' " + ", " +
                  " Телефон = '" + Telephone + "' " + 
                  " WHERE Код_клиента = " + KodKlienta;

            myCommand = new OleDbCommand(sq1, connection);
           
            connection.Open();
            myCommand.ExecuteNonQuery(); 
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM Клиенты";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox2.Visible = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";

            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox5.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox6.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox7.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox8.Text = dataGridView1[0, selectrow].Value.ToString();

            }


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";

            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox5.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox6.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox7.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox8.Text = dataGridView1[0, selectrow].Value.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            this.Height = 310;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                     "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            string FIO = textBox1.Text.ToString();
            string Adres = textBox2.Text.ToString();
            string Telephone = textBox3.Text.ToString();

            string sq1 = " INSERT INTO Клиенты " +
                  " (ФИО, Адрес, Телефон)" +
                  "  VALUES ( " +
                  " '" + FIO + "', " +
                  "  '" + Adres + "', " +
                  " '" + Telephone + "' " +
                  " )";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM Клиенты";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox1.Visible = false;
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = true;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer3.Visible = false;
            button9.Visible = false;
            groupBox2.Left = 12;
            groupBox2.Top = 284;
            this.Height = 480;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1)
            {
                MessageBox.Show("Выделите в сетке строку для удаления ");
                return;
            }

            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;

            string KodKlienta = textBox8.Text.ToString();

           
            string sq1 = "DELETE * FROM Клиенты WHERE Код_клиента = " + KodKlienta;
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM Клиенты";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox3.Visible = false;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox8.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox4.Visible = false;
            groupBox3.Visible = true;
            crystalReportViewer3.Visible = false;
            button9.Visible = false;
            groupBox3.Left = 12;
            groupBox3.Top = 284;
            this.Height = 480;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            this.Height = 310;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            /*if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выбрать критерий поиска");
                return;
            }*/

            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sql;
            sql = "SELECT * FROM Клиенты WHERE (1=1) ";

            {
                if (checkBox1.Checked)
                    sql += " AND Код_клиента = " + textBox9.Text.ToString();
            }

            {
                if (checkBox2.Checked)
                    sql += " AND ФИО Like  '%" + textBox10.Text.ToString() +"%' ";
            }

            {
                if (checkBox3.Checked)
                    sql += " AND Адрес Like '%" + textBox11.Text.ToString() + "%' ";
            }

            {
                if (checkBox4.Checked)
                    sql += " AND Телефон Like '%" + textBox12.Text.ToString() + "%' ";
            }



            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


        }

        private void button8_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT * FROM Клиенты";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
            this.Height = 310;

            checkBox1.Checked = false;
            //...
            textBox9.Text = "";
            //...
        }

        private void поискДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            textBox9.Text = "";
            groupBox4.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            //crystalReportViewer3.Visible = false;
            button9.Visible = false;
            groupBox3.Left = 12;
            groupBox3.Top = 284;
            this.Height = 435;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT * FROM Клиенты";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            ds.WriteXmlSchema("schema3.xml");

            CrystalReport3 rpt = new CrystalReport3();
            rpt.SetDataSource(ds);
            crystalReportViewer3.ReportSource = rpt;

            connection.Close();
        }

        private void отчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer3.Visible = true;
            button9.Visible = true;
            groupBox3.Left = 12;
            groupBox3.Top = 284;
            this.Height = 900;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            crystalReportViewer3.Visible = false;
            this.Height = 310;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
                textBox9.Text = "";
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
                textBox10.Text = "";
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
                textBox11.Text = "";
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
                textBox12.Text = "";
        }
    }
}

