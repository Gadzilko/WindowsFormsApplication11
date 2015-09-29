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
    public partial class Form5 : Form
    {
        public int selectrow = -1;
        public Form5()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            string KodTovara = textBox1.Text.ToString();
            string KodKlienta = textBox2.Text.ToString();
            string KolichestvoZakaza = numericUpDown1.Value.ToString();
            string Data = textBox4.Text.ToString();

            string kod_tovara = textBox1.Text.ToString();
            string kod_klienta = textBox2.Text.ToString();
            string kol_vo = numericUpDown1.Value.ToString();
            string data = dateTimePicker1.Value.ToString(); 

            string sq1 = " INSERT INTO Продажи " +
                  " (Код_товара, Код_клиента, Количество_заказа, Дата)" +
                  "  VALUES ( " +
                  " " + kod_tovara + ", " +
                  "  " + kod_klienta + ", " +
                  " " + kol_vo + ", " +
                  " '" + data + "' " + 
                  " )"; 
            
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox1.Visible = false;
        }
            

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox4.Text = dateTimePicker1.Value.ToString("dd.MM.yyyy");
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "0";
            textBox16.Text = "0";
            numericUpDown1.Value = 0;
            textBox4.Text = "";
            groupBox3.Visible = false;
            groupBox2.Visible = false;
            groupBox1.Visible = true;
            groupBox4.Visible = false;
            crystalReportViewer4.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            groupBox1.Left = 35;
            groupBox1.Top = 290;
            this.Height = 590;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            this.Height = 300;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            /*textBox8.Text = "";*/
            textBox9.Text = "";
            textBox10.Text = "";
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox5.Text = dataGridView1 [0, selectrow].Value.ToString();
                textBox6.Text = dataGridView1 [1, selectrow].Value.ToString();
                textBox7.Text = dataGridView1 [6, selectrow].Value.ToString();
               
                textBox9.Text = dataGridView1 [8, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[0, selectrow].Value.ToString();
                
                textBox17.Text = dataGridView1[3, selectrow].Value.ToString();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            /*textBox8.Text = "";*/
            textBox9.Text = "";
            textBox10.Text = "";
            
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox5.Text = dataGridView1 [0, selectrow].Value.ToString();
                textBox6.Text = dataGridView1 [1, selectrow].Value.ToString();
                textBox7.Text = dataGridView1 [6, selectrow].Value.ToString();
                
                textBox9.Text = dataGridView1 [8, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[0, selectrow].Value.ToString();
                
                textBox17.Text = dataGridView1[3, selectrow].Value.ToString();
                
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
            OleDbCommand myCommand;

            string KodProdazhi=        textBox5.Text.ToString();
            string KodTovara =         textBox6.Text.ToString();
            string KodKlienta =        textBox7.Text.ToString();
            /*string KolichestvoZakaza = textBox8.Text.ToString();*/
            string Data =              textBox9.Text.ToString();

            string sq1 = " UPDATE Продажи SET " +
                  " Код_товара = " + KodTovara + " " + ", " +
                  " Код_клиента = " + KodKlienta + " " + ", " +
                  /*" Количество_заказа = '" + KolichestvoZakaza + "', " +*/
                  " Дата = '" + Data + "' " +
                  " WHERE Код_продажи = " + KodProdazhi;

            myCommand = new OleDbCommand(sq1, connection);

            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox2.Visible = false;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            textBox9.Text = dateTimePicker1.Value.ToString("dd.MM.yyyy");
        }

        private void Form5_Load_1(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                     "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1 = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM [Каталог]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            da = new OleDbDataAdapter(myCommand);
            ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView2.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                comboBox3.Items.Add(dataGridView2[1, i].Value.ToString());
                comboBox2.Items.Add(dataGridView2[1, i].Value.ToString());
                comboBox1.Items.Add(dataGridView1[7, i].Value.ToString());
                comboBox4.Items.Add(dataGridView1[7, i].Value.ToString());
                  
            }

            if (ds.Tables["Результат"].Rows.Count > 0)
            {
                dataGridView1.Columns[0].Width = 80;
                dataGridView1.Columns[1].Width = 80;
                dataGridView1.Columns[2].Width = 80;
            }
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            /*textBox8.Text = "";*/
            textBox9.Text = "";
            groupBox3.Visible = false;
            groupBox1.Visible = false;
            groupBox2.Visible = true;
            groupBox4.Visible = false;
            crystalReportViewer4.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            groupBox2.Left = 35;
            groupBox2.Top = 290;
            this.Height = 590;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            this.Height = 300;
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
            OleDbCommand myCommand;

            string KodProdazhi = textBox10.Text.ToString();


            string sq1 = "DELETE * FROM Продажи WHERE Код_продажи = " + KodProdazhi; 
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            myCommand.ExecuteNonQuery();
            connection.Close();


            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента";
            
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
            textBox10.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = true;
            groupBox4.Visible = false;
            crystalReportViewer4.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            groupBox3.Left = 35;
            groupBox3.Top = 248;
            this.Height = 400;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            this.Height = 300;
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
            sql = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента" +
                  " WHERE (1=1) ";
            {
                if (checkBox1.Checked)
                    sql += " AND Код_продажи = " + textBox11.Text.ToString();
            }
            {
                if (checkBox2.Checked)
                    sql += " AND Код_товара = " + textBox12.Text.ToString();
            }
            {
                if (checkBox3.Checked)
                    sql += " AND Код_клиента = " + textBox13.Text.ToString();
            }
            {
                if (checkBox4.Checked)
                    sql += " AND Количество_заказа = " + textBox14.Text.ToString();
            }
            {
                if (checkBox5.Checked)
                    sql += " AND Дата = " + textBox15.Text.ToString();
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

            sq1 = " SELECT Продажи.Код_продажи, Продажи.Код_товара, Каталог.Название, Каталог.Цена, Продажи.Количество_заказа,  Каталог.Цена*Продажи.Количество_заказа AS Сумма, Продажи.Код_клиента, Клиенты.ФИО, Продажи.Дата FROM Продажи, Каталог, Клиенты WHERE Продажи.Код_товара = Каталог.Код_товара AND Продажи.Код_клиента = Клиенты.Код_клиента";
            
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
            this.Height = 300;
        }

        private void поискДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            textBox9.Text = "";
            groupBox4.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            crystalReportViewer4.Visible = false;
            button9.Visible = false;
            button10.Visible = false;
            groupBox2.Left = 35;
            groupBox2.Top = 248;
            this.Height = 600;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT Продажи.Код_продажи, Продажи.Код_товара, Продажи.Код_клиента, Клиенты.ФИО, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Цена FROM Продажи, Клиенты, Каталог WHERE Продажи.Код_клиента = Клиенты.Код_клиента AND Каталог.Код_товара = Продажи.Код_товара ";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            ds.WriteXmlSchema("schema5.xml");

            CrystalReport4 rpt = new CrystalReport4();
            rpt.SetDataSource(ds);
            crystalReportViewer4.ReportSource = rpt;

            connection.Close();
        }

        private void отчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            crystalReportViewer4.Visible = true;
            button9.Visible = true;
            button10.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox2.Left = 35;
            groupBox2.Top = 248;
            this.Height = 800;

        }

        private void button10_Click(object sender, EventArgs e)
        {
            crystalReportViewer4.Visible = false;
            this.Height = 300;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox3.SelectedIndex;
            if (r > -1)
            {
                textBox1.Text = dataGridView2[0, r].Value.ToString();
                textBox3.Text = dataGridView2[5, r].Value.ToString();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox2.SelectedIndex;
            if (r > -1)
            {
                textBox6.Text = dataGridView2[0, r].Value.ToString();
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
                textBox11.Text = "";
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
                textBox12.Text = "";
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
                textBox13.Text = "";
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
                textBox14.Text = "";
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == false)
                textBox15.Text = "";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1 = "SELECT Продажи.Код_продажи, Продажи.Код_товара FROM Продажи, Каталог WHERE Продажи.Код_товара = Каталог.Код_товара ";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();

            this.Text = sq1;

            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            textBox16.Text = Convert.ToString(Convert.ToDouble(textBox3.Text) * Convert.ToDouble(numericUpDown1.Value));
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox1.SelectedIndex;
            if (r > -1)
            {
                textBox2.Text = dataGridView1[6, r].Value.ToString();
                
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            textBox18.Text = Convert.ToString(Convert.ToDouble(textBox17.Text) * Convert.ToDouble(numericUpDown2.Value));
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox4.SelectedIndex;
            if (r > -1)
            {
                textBox7.Text = dataGridView1[6, r].Value.ToString();

            }
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            if (selectrow == -1 || selectrow >= dataGridView1.RowCount - 1)
            {
                MessageBox.Show("Выделите в сетке строку для формирования чека");
                return;
            }
        }

        private void чекToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (selectrow < (dataGridView1.RowCount - 1))
            {

                string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                        "Data Source=Kyrsovaya.mdb";
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = ConnectionString;
                OleDbCommand myCommand;
                string sq1;

                sq1 = "SELECT * FROM [Вид техники]"; ///////////
                myCommand = new OleDbCommand(sq1, connection);
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
                DataSet ds = new DataSet();
                da.Fill(ds, "Результат");
                ds.WriteXmlSchema("schema.xml"); ///////

                CrystalReport1 rpt = new CrystalReport1();////////////

                rpt.SetDataSource(ds);
                crystalReportViewer1.ReportSource = rpt;
                connection.Close();

                crystalReportViewer1.Visible = true;
            }
        }
    }
}
