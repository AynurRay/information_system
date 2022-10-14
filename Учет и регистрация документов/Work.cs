using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace Учет_и_регистрация_документов
{
    public partial class Work : Form
    {
        public Work()
        {
            InitializeComponent();
        }

        private SQLiteConnection DB;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Work_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection("Data Source = DocDB.db; Version=3");
            DB.Open();

            All();
        }


        public void All()
        {
            DataSet ds = new DataSet(); //Создаем объект класса DataSet

            string sql = "SELECT * FROM worker"; //Sql запрос (достать все из таблицы customer)

            string path = "DocDB.db"; //Путь к файлу БД

            string ConnectionString = "Data Source=" + path + ";Version=3;New=True;Compress=True;"; //Строка соеденения (так хочет sqlite)

            SQLiteConnection conn = new SQLiteConnection(ConnectionString); //Создаем соеденение

            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, conn);//Создаем объект класса DataAdapter (тут мы передаем наш запрос и получаем ответ)

            da.Fill(ds);//Заполняем DataSet cодержимым DataAdapter'a

            dataGridView1.DataSource = ds.Tables[0].DefaultView;//Заполняем созданный на форме dataGridView1

            conn.Close();//Закрываем соединение

            dataGridView1.AllowUserToAddRows = false;
        }


        private void button2_Click(object sender, EventArgs e)
        {

            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && comboBox1.Text!="")
            {


                try
                {
                    string fio = textBox1.Text + " " + textBox2.Text + " " + textBox3.Text + " ";

                    SQLiteCommand CMD = DB.CreateCommand();
                    CMD.CommandText = "INSERT INTO worker (ФИО, Должность) VALUES (@ФИО, @Должность)";
                    CMD.Parameters.Add("@ФИО", System.Data.DbType.String).Value = fio.ToUpper();
                    CMD.Parameters.Add("@Должность", System.Data.DbType.String).Value = comboBox1.Text.ToUpper();

                    CMD.ExecuteNonQuery();

                    MessageBox.Show("Данные успешно добавлены!");
                    All();
                }

                catch { MessageBox.Show("Ошибка при добавлении!"); }
            }
        }

        private void button3_Click(object sender, EventArgs e)
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
                comand.CommandText = "DELETE FROM worker WHERE Индекс='" + id + "'";
                comand.ExecuteNonQuery();

                All();

                MessageBox.Show("Данные успешно удалены!");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                int index = dataGridView1.SelectedCells[0].RowIndex;

                textBox7.Text = Convert.ToString(dataGridView1.Rows[index].Cells[0].Value);
                string fio = Convert.ToString(dataGridView1.Rows[index].Cells[1].Value);

                string[] a = fio.Split(' ');

                textBox6.Text = a[0];
                textBox5.Text = a[1];
                textBox4.Text = a[2];
                comboBox2.Text = Convert.ToString(dataGridView1.Rows[index].Cells[2].Value);
            }

            catch (IndexOutOfRangeException)
            {

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox6.Text != "" && textBox5.Text != "" && textBox4.Text != "" && comboBox2.Text != "")
            {
                try
                {
                    string fio = textBox6.Text + " " + textBox5.Text + " " + textBox4.Text + " ";

                    SQLiteCommand CMD = DB.CreateCommand();

                    SQLiteCommand comand = DB.CreateCommand();
                    comand.CommandText = "UPDATE worker SET ФИО='" + fio.ToUpper() + "', Должность='" + comboBox2.Text.ToUpper() + "' WHERE Индекс='" + textBox7.Text + "'";
                    comand.ExecuteNonQuery();

                    All();

                    MessageBox.Show("Данные успешно обновлены!");
                }

                catch
                {
                    MessageBox.Show("При редактировании возникла ошибка!");
                }
            }
        }
    }
}
