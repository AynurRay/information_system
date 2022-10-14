using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SQLite;

using Excel = Microsoft.Office.Interop.Excel;

namespace Учет_и_регистрация_документов
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /*Создаем подключение*/
        private SQLiteConnection DB;

        private void Form1_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection("Data Source = DocDB.db; Version=3");
            DB.Open();

            Timer timer1 = new Timer();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 100;
            timer1.Start();

            All();


            panel2.BackColor = Color.LightGray;


            SQLiteCommand command = new SQLiteCommand("select ФИО from worker", DB);

            SQLiteDataReader reader = command.ExecuteReader();
            while (reader.Read())
                fio.Items.Add((string)reader["ФИО"]);

        }

        // Создаём объекты Image с нашими картинками, при чём путь указан в директорию, где лежит наш .exe
        Image image1 = Image.FromFile(Environment.CurrentDirectory + @"\" + "3.png");
        Image image2 = Image.FromFile(Environment.CurrentDirectory + @"\" + "3_.png");



        private void exit_Box_MouseMove(object sender, MouseEventArgs e)
        {
            exit_Box.BackgroundImage = image1;
        }

        private void exit_Box_MouseLeave(object sender, EventArgs e)
        {
            exit_Box.BackgroundImage = image2;
        }

        /*Удалить*/
        private void Delete_Click(object sender, EventArgs e)
        {
            var result = new System.Windows.Forms.DialogResult();
            result = MessageBox.Show("Удалить данные?", "Удалить",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                /*обработчик*/
            }
        }

        /*Выход*/
        private void exit_Box_Click(object sender, EventArgs e)
        {
            var result = new System.Windows.Forms.DialogResult();
            result = MessageBox.Show("Завершить работу программы?", "Выход",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }

        }

      
        private void Form1_Activated(object sender, EventArgs e)
        {
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime ThToday = DateTime.Now;
            string ThData = ThToday.ToString("Текущая дата: " + "dd MMMM yyyy" + " г." + " Время: " + "HH:mm:ss");
            this.label1.Text = ThData;
        }

        Image image3 = Image.FromFile(Environment.CurrentDirectory + @"\" + "2.png");
        Image image4 = Image.FromFile(Environment.CurrentDirectory + @"\" + "2_.png");

        private void pictureBox2_MouseLeave(object sender, EventArgs e)
        {
            pictureBox2.BackgroundImage = image4;
        }

        private void pictureBox2_MouseMove(object sender, MouseEventArgs e)
        {
            pictureBox2.BackgroundImage = image3;
        }

        Image image5 = Image.FromFile(Environment.CurrentDirectory + @"\" + "1.png");
        Image image6 = Image.FromFile(Environment.CurrentDirectory + @"\" + "1_.png");


        private void pictureBox3_MouseLeave(object sender, EventArgs e)
        {
            pictureBox3.BackgroundImage = image6;
        }

        private void pictureBox3_MouseMove(object sender, MouseEventArgs e)
        {
            pictureBox3.BackgroundImage = image5;
        }

        About about;
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (about == null || about.IsDisposed)
            {
                about = new About();
                about.Show();
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string commandText = "Руководство пользователя.pdf";
            var proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = commandText;
            proc.StartInfo.UseShellExecute = true;
            proc.Start();
        }

        private void button3_Click(object sender, EventArgs e)
        {
       
        }

        Add add;
        private void button5_Click(object sender, EventArgs e)
        {
            if (add == null || add.IsDisposed)
            {
                add = new Add();
                add.Show();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var result = new System.Windows.Forms.DialogResult();
            result = MessageBox.Show("Удалить данные?", "Удалить",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {

                int index = dataGridView1.SelectedCells[0].RowIndex;

                string id = Convert.ToString(dataGridView1.Rows[index].Cells[0].Value);//id 

                SQLiteCommand CMD = DB.CreateCommand();


                SQLiteCommand comand = DB.CreateCommand();
                comand.CommandText = "DELETE FROM doc WHERE Код='" + id + "'";
                comand.ExecuteNonQuery();

                All();

                MessageBox.Show("Данные успешно удалены!");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var result = new System.Windows.Forms.DialogResult();
            result = MessageBox.Show("Внимание! Вы собираетесь удалить все данные! Продолжить?", "Удалить все данные",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                SQLiteCommand CMD = DB.CreateCommand();

                SQLiteCommand comand = DB.CreateCommand();
                comand.CommandText = "DELETE FROM doc";
                comand.ExecuteNonQuery();

                All();

                MessageBox.Show("Данные успешно удалены!");
            }
        }

        private void nositel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (nositel.Text == "БУМАЖНЫЙ НОСИТЕЛЬ")
            {
                stellaj.Enabled = true;
                polka.Enabled = true;
                papka.Enabled = true;

                mesto.Enabled = false;
                button11.Enabled = false;

                mesto.Text = "";
            }

            if (nositel.Text == "ЭЛЕКТРОННЫЙ ДОКУМЕНТ")
            {
                stellaj.Enabled = false;
                polka.Enabled = false;
                papka.Enabled = false;

                mesto.Enabled = true;
                button11.Enabled = true;

                stellaj.Text = "";
                polka.Text = "";
                papka.Text = "";

                mesto.Text = "";
            }

            if (nositel.Text == "")
            {
                stellaj.Enabled = false;
                polka.Enabled = false;
                papka.Enabled = false;

                mesto.Enabled = false;
                button11.Enabled = false;

                stellaj.Text = "";
                polka.Text = "";
                papka.Text = "";

                mesto.Text = "";

            }

            World();
        }

        public void All_main()
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Тип=='ПРИЕМ' OR Тип=='ПЕРЕВОД' OR Тип=='УВОЛНЕНИЕ'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView2.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение

        }



        public void All_others()
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Тип!='ПРИЕМ' AND Тип!='ПЕРЕВОД' AND Тип!='УВОЛНЕНИЕ'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView3.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение

        }

        public void All()
        {

            try
            {
                DataSet ds = new DataSet(); //Создаем объект класса DataSet

                string sql = "SELECT * FROM doc"; //Sql запрос (достать все из таблицы customer)

                string path = "DocDB.db"; //Путь к файлу БД

                string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

                SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

                SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

                da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

                dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

                conn.Close();//Закрываем соединение

                All_main();
                All_others();

                Color_();

                dataGridView4.Rows.Clear();
                dataGridView4.Rows.Add("№ Документа", "Тип документа", "Носитель", "Наименование документа", "Дата", "ФИО сотрудника", "Должность", "Комментарий", "Отметка об исполнении", "Местонахождение оригинала");

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView4.Rows.Add(dataGridView1.Rows[i].Cells[1].Value, dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value, dataGridView1.Rows[i].Cells[5].Value, dataGridView1.Rows[i].Cells[6].Value, dataGridView1.Rows[i].Cells[7].Value, dataGridView1.Rows[i].Cells[8].Value, dataGridView1.Rows[i].Cells[9].Value, dataGridView1.Rows[i].Cells[10].Value);

                }
            }

            catch (ArgumentOutOfRangeException) { }

        }

        public void Color_()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                for (int j = i + 1; j < dataGridView1.Rows.Count; j++)
                    if (Convert.ToString(dataGridView1.Rows[i].Cells[9].Value) == "ИСПОЛНЕНО")
                    {
                        dataGridView1.Rows[j - 1].DefaultCellStyle.BackColor = Color.White;
                    }

                    else dataGridView1.Rows[j - 1].DefaultCellStyle.BackColor = Color.LightGray;
        }


        public void World()
        {
            if (nomer.Text == "" || tip.Text == "" || nositel.Text == "" || date.Text == "" || fio.Text == "" || doljnost.Text == "" || otmetka.Text == "" || mesto.Text == "" || name.Text == "" || coment.Text == "")
            {
                button12.Enabled = false;
            }

            else
                button12.Enabled = true;
        }


        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime dt = date.Value;

                SQLiteCommand CMD = DB.CreateCommand();

                SQLiteCommand comand = DB.CreateCommand();
                comand.CommandText = "UPDATE doc SET №_Документа='" + nomer.Text.ToUpper() + "', Тип='" + tip.Text.ToUpper() + "', Носитель='" + nositel.Text.ToUpper() + "', Наименование_документа='" + name.Text.ToUpper() + "', Дата='" + dt.ToString("dd.MM.yyyy") + "', ФИО='" + fio.Text.ToUpper() + "', Должность='" + doljnost.Text.ToUpper() + "', Комментарий='" + coment.Text.ToUpper() + "', Отметка_об_исполнении='" + otmetka.Text.ToUpper() + "', Местонахождение_оригинала='" + mesto.Text.ToUpper() + "' WHERE №_Документа='" + nomer.Text + "'";
                comand.ExecuteNonQuery();

                All();

                MessageBox.Show("Данные успешно обновлены!");
            }

            catch
            {
                MessageBox.Show("При редактировании возникла ошибка!");
            }
        }

        private void nomer_TextChanged(object sender, EventArgs e)
        {
            World();
        }

        private void tip_SelectedIndexChanged(object sender, EventArgs e)
        {
            World();
        }

        private void doljnost_SelectedIndexChanged(object sender, EventArgs e)
        {
            World();
        }

        private void otmetka_SelectedIndexChanged(object sender, EventArgs e)
        {
            World();
        }

        private void mesto_TextChanged(object sender, EventArgs e)
        {
            World();
        }

        private void papka_TextChanged(object sender, EventArgs e)
        {
            if (papka.Text != "" && stellaj.Text != "" && polka.Text != "")
            {
                mesto.Text = stellaj.Text + "/" + polka.Text + "/" + papka.Text;
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            All();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView1.SelectedCells[0].RowIndex;
            string nositel = Convert.ToString(dataGridView1.Rows[index].Cells[3].Value);

            if (nositel == "ЭЛЕКТРОННЫЙ ДОКУМЕНТ")
            {
                button8.Enabled = true;
            }

            else button8.Enabled = false;

            nomer.Text = Convert.ToString(dataGridView1.Rows[index].Cells[1].Value);
            tip.Text = Convert.ToString(dataGridView1.Rows[index].Cells[2].Value);
            this.nositel.Text = Convert.ToString(dataGridView1.Rows[index].Cells[3].Value);
            name.Text = Convert.ToString(dataGridView1.Rows[index].Cells[4].Value);
            date.Text = Convert.ToString(dataGridView1.Rows[index].Cells[5].Value);
            fio.Text = Convert.ToString(dataGridView1.Rows[index].Cells[6].Value);
            doljnost.Text = Convert.ToString(dataGridView1.Rows[index].Cells[7].Value);
            coment.Text = Convert.ToString(dataGridView1.Rows[index].Cells[8].Value);
            otmetka.Text = Convert.ToString(dataGridView1.Rows[index].Cells[9].Value);
            mesto.Text = Convert.ToString(dataGridView1.Rows[index].Cells[10].Value);

            if (nositel == "БУМАЖНЫЙ НОСИТЕЛЬ")
            {
                string[] words = mesto.Text.Split('/');
                stellaj.Text = words[0].ToString();
                polka.Text = words[1].ToString();
                papka.Text = words[2].ToString();
            }

            else
            {
                stellaj.Text = "";
                polka.Text = "";
                papka.Clear();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                int index = dataGridView1.SelectedCells[0].RowIndex;
                string nositel = Convert.ToString(dataGridView1.Rows[index].Cells[3].Value);


                if (nositel == "ЭЛЕКТРОННЫЙ ДОКУМЕНТ")
                {
                    string commandText = Convert.ToString(dataGridView1.Rows[index].Cells[10].Value);
                    var proc = new System.Diagnostics.Process();
                    proc.StartInfo.FileName = commandText;
                    proc.StartInfo.UseShellExecute = true;
                    proc.Start();
                }

            }

            catch { MessageBox.Show("Файл " + Path.GetFileName(mesto.Text) + " не найден!"); }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string str = dialog.FileName;
                mesto.Text = str;
            }
        }

        private void name_TextChanged(object sender, EventArgs e)
        {
            World();
        }

        private void coment_TextChanged(object sender, EventArgs e)
        {
            World();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE ФИО LIKE '%" + textBox1.Text.ToUpper() + "%'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Наименование_документа LIKE '%" + textBox4.Text.ToUpper() + "%'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }


        Work work;
        private void button16_Click(object sender, EventArgs e)
        {
            if (work == null || work.IsDisposed)
            {
                work = new Work();
                work.Show();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Тип='" + comboBox1.Text.ToUpper() + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Носитель='" + comboBox3.Text.ToUpper() + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Отметка_об_исполнении='" + comboBox2.Text.ToUpper() + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DateTime dt = dateTimePicker3.Value;

            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Дата='" + dt.ToString("dd.MM.yyyy") + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение


            Color_();
        }

  
        statistic statistic;
        private void button9_Click(object sender, EventArgs e)
        {
            if (statistic == null || statistic.IsDisposed)
            {
                statistic = new statistic();
                statistic.Show();
            }
        }

        private void fio_SelectedIndexChanged(object sender, EventArgs e)
        {
            World();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Excel (*.xls)|*.xls|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excelapp = new Excel.Application();
                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                for (int i = 1; i < dataGridView4.RowCount + 1; i++)
                {
                    for (int j = 1; j < dataGridView4.ColumnCount + 1; j++)
                    {
                        worksheet.Rows[i].Columns[j] = dataGridView4.Rows[i - 1].Cells[j - 1].Value;
                    }
                }

                excelapp.AlertBeforeOverwriting = false;
                //workbook.SaveAs(@"C:\Users\1\Desktop\файл.xls");
                excelapp.Quit();
            }

        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Тип=='" + comboBox6.Text.ToUpper() + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView2.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }

        private void button17_Click_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Носитель=='" + comboBox4.Text.ToUpper() + "'AND (Тип=='ПРИЕМ' OR Тип=='ПЕРЕВОД' OR Тип=='УВОЛНЕНИЕ')"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView2.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE ФИО LIKE'%" + textBox2.Text.ToUpper() + "%'AND (Тип=='ПРИЕМ' OR Тип=='ПЕРЕВОД' OR Тип=='УВОЛНЕНИЕ')"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView2.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }

        private void button21_Click_1(object sender, EventArgs e)
        {

            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Тип=='" + comboBox9.Text.ToUpper() + "'"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView3.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }

        private void button20_Click_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE Носитель=='" + comboBox7.Text.ToUpper() + "' AND (Тип!='ПРИЕМ' AND Тип!='ПЕРЕВОД' AND Тип!='УВОЛНЕНИЕ')"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView3.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM doc WHERE ФИО LIKE'%" + textBox3.Text.ToUpper() + "%' AND (Тип!='ПРИЕМ' AND Тип!='ПЕРЕВОД' AND Тип!='УВОЛНЕНИЕ')"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView3.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение
        }
    }
}
