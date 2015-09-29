using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public int selectrow = -1;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1 = "SELECT * FROM [Вид техники]";
            OleDbCommand myCommand = new OleDbCommand (sq1, connection);
            connection.Open();

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            if (ds.Tables["Результат"].Rows.Count > 0)
            {
                dataGridView1.Columns[0].Width = 70;
                dataGridView1.Columns[1].Width = 180;
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1;
            string Nazvanie = textBox1.Text.ToString();
            sq1 = "INSERT INTO [Вид техники] ( Название )" +
                  " VALUES (" +
                  "'" + Nazvanie + "' " +
                  ")";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            groupBox1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            this.Height = 240;
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox1.Visible = true;
            groupBox4.Visible = false;
            crystalReportViewer1.Visible = false;
            button9.Visible = false;
            groupBox1.Left = 12;
            groupBox1.Top = 216;
            this.Height = 365;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox2.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox3.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox2.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox3.Text = dataGridView1[1, selectrow].Value.ToString();
                textBox4.Text = dataGridView1[0, selectrow].Value.ToString();
            }
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

            string sq1;
            string Nazvanie = textBox3.Text.ToString();
            string KodVida = textBox2.Text.ToString();

            sq1 = "UPDATE [Вид техники] SET " +
                  "Название =  '" + Nazvanie + "' " +
                  " WHERE Код_вида = " + KodVida;
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            groupBox2.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            this.Height = 240;
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";

            groupBox1.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer1.Visible = false;
            button9.Visible = false;
            groupBox2.Left = 12;
            groupBox2.Top = 216;
            this.Height = 365;

            groupBox2.Visible = true;
        }

        private void button5_Click(object sender, EventArgs e)
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

            string sq1;
            string KodVida = textBox4.Text.ToString();

            sq1 = "DELETE * FROM [Вид техники] WHERE Код_вида = " + KodVida;
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            groupBox3.Visible = false;
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            this.Height = 240;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox4.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox4.Visible = false;
            groupBox3.Visible = true;
            crystalReportViewer1.Visible = false;
            button9.Visible = false;
            groupBox3.Left = 12;
            groupBox3.Top = 216;
            this.Height = 365;
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
            sql = "SELECT * FROM [Вид техники] WHERE (1=1) ";

            {
                if (checkBox1.Checked)
                    sql += " AND Код_вида = " + textBox5.Text.ToString();
            }

            {
                if (checkBox2.Checked)
                    sql += " AND Название Like '%" + textBox6.Text.ToString() +"%' ";
            }



            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            groupBox1.Visible = false;
            this.Text = sql;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
            this.Height = 240;
        }

        private void поискДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            textBox5.Text = "";
            groupBox4.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            crystalReportViewer1.Visible = false;
            button9.Visible = false;
            //groupBox4.Left = 12;
            //groupBox4.Top = 216;
            this.Height = 300;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                     "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            ds.WriteXmlSchema("schema.xml");

            CrystalReport1 rpt = new CrystalReport1();
            rpt.SetDataSource(ds);
            crystalReportViewer1.ReportSource = rpt;

            connection.Close();
        }

        private void отчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer1.Visible = true;
            button9.Visible = true;
            groupBox3.Left = 12;
            groupBox3.Top = 216;
            this.Height = 900;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            crystalReportViewer1.Visible = false;
            this.Height = 240;
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
                textBox5.Text = "";
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
                textBox6.Text = "";
        }

      
    }
}
