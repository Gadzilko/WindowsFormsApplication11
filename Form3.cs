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
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public int selectrow = -1;
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string sq1 = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида";
            OleDbCommand myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            if (dataGridView1.RowCount > 0)
            {
                if (!dataGridView1[7, 0].Value.ToString().Equals("") && File.Exists(@"img/" + dataGridView1[7, 0].Value.ToString() + ".jpg"))
                    pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, 0].Value.ToString() + ".jpg");
                else
                    pictureBox1.Image = new Bitmap(@"img/dom.jpg");
               // pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, 0].Value.ToString() + ".jpg");
            }


            //запрос для подстановки данных о виде техники
            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sq1 = "SELECT * FROM [Вид техники]";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            da = new OleDbDataAdapter(myCommand);
            ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView2.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();
            for (int i = 0; i < dataGridView2.RowCount - 1; i++)
            {
                comboBox2.Items.Add(dataGridView2[1,i].Value.ToString());
                comboBox3.Items.Add(dataGridView2[1, i].Value.ToString());

                
            }
            if (ds.Tables["Результат"].Rows.Count > 0)
            {
                dataGridView1.Columns[0].Width = 60;
                dataGridView1.Columns[1].Width = 180;
                dataGridView1.Columns[2].Width = 70;
                dataGridView1.Columns[3].Width = 170;
                dataGridView1.Columns[4].Width = 130;
                dataGridView1.Columns[5].Width = 170;
                dataGridView1.Columns[6].Width = 60;
                dataGridView1.Columns[7].Width = 60;
                dataGridView1.Columns[8].Width = 60;
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;

            string Nazvanie = textBox1.Text.ToString();
            string KodVida = textBox2.Text.ToString();
            string Firma = textBox3.Text.ToString();
            string Xaraktiristiki = textBox4.Text.ToString();
            string Cena = textBox5.Text.ToString();
            string Foto = textBox6.Text.ToString();
            string Garantiya = textBox7.Text.ToString();
            string sql = " INSERT INTO Каталог " +
                  " (Название, Код_вида, Фирма, Характиристики, Цена, Фото, Гарантия)" +
                  "  VALUES ( " +
                  " '" + Nazvanie + "', " +
                  "  " + KodVida + ", " +
                  " '" + Firma + "', " +
                  " '" + Xaraktiristiki + "', " +
                  Cena + ", " +
                  " '" + Foto + "', " +
                  " '" + Garantiya + "' " +
                  " )";
            
                OleDbCommand myCommand = new OleDbCommand(sql, connection);
            try
            {
                connection.Open();
                myCommand.ExecuteNonQuery();
                connection.Close();
            }
            catch { MessageBox.Show("Ошибка добавления данных"); }

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида";
                            
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox1.Visible = false;
        }

      
        private void добавитьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox1.Visible = true;
            crystalReportViewer2.Visible = false;
            button9.Visible = false;
            groupBox1.Left = 35;
            groupBox1.Top = 248;
            /*this.Height = 470;*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            this.Height = 280;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";

            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox8.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox9.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox11.Text = dataGridView1[4, selectrow].Value.ToString();
                textBox12.Text = dataGridView1[5, selectrow].Value.ToString();
                textBox13.Text = dataGridView1[6, selectrow].Value.ToString();
                textBox14.Text = dataGridView1[7, selectrow].Value.ToString();
                textBox15.Text = dataGridView1[8, selectrow].Value.ToString();
                textBox16.Text = dataGridView1[2, selectrow].Value.ToString();
                //pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg");
                if (dataGridView1[7, selectrow].Value.ToString() != "" && File.Exists(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg"))
                    pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg");
                else
                    pictureBox1.Image = new Bitmap(@"img/dom.jpg");

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            selectrow = dataGridView1.CurrentCell.RowIndex;
            if (selectrow < (dataGridView1.RowCount - 1))
            {
                textBox8.Text = dataGridView1[2, selectrow].Value.ToString();
                textBox9.Text = dataGridView1[3, selectrow].Value.ToString();
                textBox10.Text = dataGridView1[0, selectrow].Value.ToString();
                textBox11.Text = dataGridView1[4, selectrow].Value.ToString();
                textBox12.Text = dataGridView1[5, selectrow].Value.ToString();
                textBox13.Text = dataGridView1[6, selectrow].Value.ToString();
                textBox14.Text = dataGridView1[7, selectrow].Value.ToString();
                textBox15.Text = dataGridView1[8, selectrow].Value.ToString();
                textBox16.Text = dataGridView1[2, selectrow].Value.ToString();
                //pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg");
                if (dataGridView1[7, selectrow].Value.ToString() != "" && File.Exists(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg"))
                    pictureBox1.Image = new Bitmap(@"img/" + dataGridView1[7, selectrow].Value.ToString() + ".jpg");
                else
                    pictureBox1.Image = new Bitmap(@"img/dom.jpg");
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

            string KodTovara = textBox8.Text.ToString();
            string Nazvanie = textBox9.Text.ToString();
            string KodVida = textBox10.Text.ToString();
            string Firma = textBox11.Text.ToString();
            string Xaraktiristiki = textBox12.Text.ToString();
            string Cena = textBox13.Text.ToString();
            string Foto = textBox14.Text.ToString();
            string Garantiya = textBox15.Text.ToString();
            string sql = " UPDATE Каталог SET " +
                  " Название = '" + Nazvanie + "' " + ", " +
                  " Код_вида = " + KodVida + ", " +
                  " Фирма = '" + Firma + "' " + ", " +
                  " Характиристики = '" + Xaraktiristiki + "' " + ", " +
                  " Цена = " + Cena + ", " +
                  " Фото = '" + Foto + "' " + ", " +
                  " Гарантия = '" + Garantiya + "' " +
                  " WHERE Код_товара = " + KodTovara;
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            try
            {
                connection.Open();
                myCommand.ExecuteNonQuery();
                connection.Close();
            }
            catch { MessageBox.Show("Ошибка редактирования данных"); }

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида";
                       
            myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();


            groupBox2.Visible = false;
        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            groupBox1.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox2.Visible = true;
            crystalReportViewer2.Visible = false;
            button9.Visible = false;
            groupBox2.Left = 35;
            groupBox2.Top = 275;
            this.Height = 550;
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

            string KodTovara = textBox16.Text.ToString();
            string sql = "DELETE * FROM Каталог WHERE Код_товара = " + KodTovara;
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            try
            {
                connection.Open();
                myCommand.ExecuteNonQuery();
                connection.Close();
            }
            catch { MessageBox.Show("Ошибка удаления данных"); }

            connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            sql = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида";
            
            myCommand = new OleDbCommand(sql, connection);
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
            textBox16.Text = "";
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox4.Visible = false;
            groupBox3.Visible = true;
            crystalReportViewer2.Visible = false;
            button9.Visible = false;
            groupBox3.Left = 35;
            groupBox3.Top = 248;
            /*this.Height = 400;*/
        }

        private void button6_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            this.Height = 280;

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
            sql = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида " +
                  " AND (1=1) ";
            {
                if (checkBox8.Checked)
                    sql += " AND Каталог.Код_товара = " + textBox24.Text.ToString();
            }
            
            {
                if (checkBox1.Checked)
                    sql += " AND [Вид техники].Код_вида = " + textBox17.Text.ToString();
            }

            {
                if (checkBox2.Checked)
                    sql += " AND Каталог.Название Like  '%" + textBox18.Text.ToString() + "%' ";
            }

            {
                if (checkBox3.Checked)
                    sql += " AND Каталог.Фирма Like  '%" + textBox19.Text.ToString() + "%' ";
            }

            {
                if (checkBox4.Checked)
                    sql += " AND Каталог.Характиристики Like  '%" + textBox20.Text.ToString() + "%'";
            }

            {
                if (checkBox5.Checked)
                    sql += " AND Каталог.Цена = " + textBox21.Text.ToString();
            }

            {
                if (checkBox6.Checked)
                    sql += " AND Каталог.Фото Like  '%" + textBox22.Text.ToString() + "%' ";
            }


            {
                if (checkBox7.Checked)
                    sql += " AND Каталог.Гарантия Like '%" + textBox23.Text.ToString() + "%' ";
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

            sq1 = "SELECT [Вид техники].Код_вида, [Вид техники].Название, Каталог.Код_товара, Каталог.Название, Каталог.Фирма, Каталог.Характиристики, Каталог.Цена, Каталог.Фото, Каталог.Гарантия FROM Каталог, [Вид техники] WHERE [Вид техники].Код_вида = Каталог.Код_вида";
            
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            dataGridView1.DataSource = ds.Tables["Результат"].DefaultView;
            connection.Close();

            groupBox4.Visible = false;
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox22.Text = "";
            textBox23.Text = "";
            textBox24.Text = "";

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;

            //this.Height = 280;

        }

        private void поискДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            textBox17.Text = "";
            groupBox4.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            crystalReportViewer2.Visible = false;
            button9.Visible = false;
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                                      "Data Source=Kyrsovaya.mdb";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = ConnectionString;
            OleDbCommand myCommand;
            string sq1;

            sq1 = "SELECT * FROM Каталог";
            myCommand = new OleDbCommand(sq1, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Результат");
            ds.WriteXmlSchema("schema2.xml");

            CrystalReport2 rpt = new CrystalReport2();
            rpt.SetDataSource(ds);
            crystalReportViewer2.ReportSource = rpt;

            connection.Close();
        }

        private void отчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            crystalReportViewer2.Visible = true;
            button9.Visible = true;
            groupBox3.Left = 12;
            groupBox3.Top = 284;
            this.Height = 900;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox2.SelectedIndex;
            if(r>-1)
            textBox2.Text = dataGridView2[0, r].Value.ToString();
            textBox5.Text = dataGridView1[6, r].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = false;
            this.Height = 280;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            crystalReportViewer2.Visible = false;
            this.Height = 280;
        }


        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
                
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int r = comboBox3.SelectedIndex;
            if (r > -1)
                textBox10.Text = dataGridView2[0, r].Value.ToString();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == false)
                textBox24.Text = "";
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
                textBox17.Text = "";
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
                textBox18.Text = "";
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
                textBox19.Text = "";
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
                textBox20.Text = "";
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == false)
                textBox21.Text = "";
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == false)
                textBox22.Text = "";
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == false)
                textBox23.Text = "";
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            textBox25.Text = Convert.ToString(Convert.ToDouble(textBox5.Text) * Convert.ToDouble(numericUpDown1.Value));
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            textBox26.Text = Convert.ToString(Convert.ToDouble(textBox13.Text) * Convert.ToDouble(numericUpDown2.Value));
        }

        

        

       
        
    }
}
