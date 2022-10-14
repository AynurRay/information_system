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
    public partial class Add : Form
    {
        public Add()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private SQLiteConnection DB;

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string str = dialog.FileName;
                mesto.Text = str;
            }
        }

        public void World()
        {
            if (nomer.Text == "" || tip.Text == "" || nositel.Text == "" || date.Text == "" || fio.Text == "" || doljnost.Text == "" || otmetka.Text == "" || mesto.Text == "" || name.Text=="" || soder.Text=="")
            {
                button4.Enabled = false;
            }

            else
                button4.Enabled = true;
        }

        private void Add_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection("Data Source = DocDB.db; Version=3");
            DB.Open();

            SQLiteCommand command = new SQLiteCommand("select ФИО from worker", DB);

            SQLiteDataReader reader = command.ExecuteReader();
            while (reader.Read())
                fio.Items.Add((string)reader["ФИО"]);         //СтолбецТаблицы

        }

        private void nomer_TextChanged(object sender, EventArgs e)
        {
            if (nomer.Text != "")
            {
                nomer.BackColor = Color.YellowGreen;
            }

            else nomer.BackColor = Color.White;

            World();

            try
            {
                Convert.ToInt64(nomer.Text);
            }

            catch (FormatException)
            {
                nomer.Clear();
            }

        }

        private void tip_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tip.Text != "")
            {
                tip.BackColor = Color.YellowGreen;
            }

            else tip.BackColor = Color.White;

            World();
        }

        private void nositel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (nositel.Text == "БУМАЖНЫЙ НОСИТЕЛЬ")
            {
                stellaj.Enabled = true;
                polka.Enabled = true;
                papka.Enabled = true;

                mesto.Enabled = false;
                button1.Enabled = false;

                mesto.Text = "";
            }

            if (nositel.Text == "ЭЛЕКТРОННЫЙ ДОКУМЕНТ")
            {
                stellaj.Enabled = false;
                polka.Enabled = false;
                papka.Enabled = false;

                mesto.Enabled = true;
                button1.Enabled = true;

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
                button1.Enabled = false;

                stellaj.Text = "";
                polka.Text = "";
                papka.Text = "";

                mesto.Text = "";

                nositel.BackColor = Color.White;
            }

            if (nositel.Text != "")
            {
                nositel.BackColor = Color.YellowGreen;
            }

            World();
        }

        private void doljnost_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (doljnost.Text != "")
            {
                doljnost.BackColor = Color.YellowGreen;
            }

            else doljnost.BackColor = Color.White;

            World();
        }

        private void otmetka_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (otmetka.Text != "")
            {
                otmetka.BackColor = Color.YellowGreen;
            }

            else otmetka.BackColor = Color.White;

            World();
        }

        private void mesto_TextChanged(object sender, EventArgs e)
        {
            if (mesto.Text != "")
            {
                mesto.BackColor = Color.YellowGreen;
            }

            else mesto.BackColor = Color.White;

            World();
        }

        private void papka_TextChanged(object sender, EventArgs e)
        {
            if (papka.Text != "" && stellaj.Text != "" && polka.Text != "")
            {
                papka.BackColor = Color.YellowGreen;

                mesto.Text = stellaj.Text + "/" + polka.Text + "/" + papka.Text;
            }

            else papka.BackColor = Color.White;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            nomer.Clear();
            tip.Text = "";
            nositel.Text = "";
            date.Text = "";
            fio.Text = "";
            doljnost.Text = "";
            otmetka.Text = "";
            mesto.Text = "";
            stellaj.Text = "";
            polka.Text = "";
            papka.Text = "";
            name.Clear();
            soder.Clear();

            fio.BackColor = Color.White;

        }

        private void stellaj_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (stellaj.Text != "")
            {
                stellaj.BackColor = Color.YellowGreen;
            }

            else stellaj.BackColor = Color.White;
        }

        private void polka_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (polka.Text != "")
            {
                polka.BackColor = Color.YellowGreen;
            }

            else polka.BackColor = Color.White;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime dt = date.Value;

                SQLiteCommand CMD = DB.CreateCommand();
                CMD.CommandText = "INSERT INTO doc (№_Документа, Тип, Носитель, Наименование_документа, Дата, ФИО, Должность, Комментарий, Отметка_об_исполнении, Местонахождение_оригинала) VALUES (@№_Документа, @Тип, @Носитель, @Наименование_документа, @Дата, @ФИО, @Должность, @Комментарий, @Отметка_об_исполнении, @Местонахождение_оригинала)";
                CMD.Parameters.Add("@№_Документа", System.Data.DbType.String).Value = nomer.Text.ToUpper();
                CMD.Parameters.Add("@Тип", System.Data.DbType.String).Value = tip.Text.ToUpper();
                CMD.Parameters.Add("@Носитель", System.Data.DbType.String).Value = nositel.Text.ToUpper();
                CMD.Parameters.Add("@Наименование_документа", System.Data.DbType.String).Value = name.Text.ToUpper();
                CMD.Parameters.Add("@Дата", System.Data.DbType.String).Value = dt.ToString("dd.MM.yyyy");
                CMD.Parameters.Add("@ФИО", System.Data.DbType.String).Value = fio.Text.ToUpper();
                CMD.Parameters.Add("@Должность", System.Data.DbType.String).Value = doljnost.Text.ToUpper();
                CMD.Parameters.Add("@Комментарий", System.Data.DbType.String).Value = soder.Text.ToUpper();
                CMD.Parameters.Add("@Отметка_об_исполнении", System.Data.DbType.String).Value = otmetka.Text.ToUpper();
                CMD.Parameters.Add("@Местонахождение_оригинала", System.Data.DbType.String).Value = mesto.Text.ToUpper();

                CMD.ExecuteNonQuery();

                MessageBox.Show("Данные успешно добавлены!");
            }

            catch { MessageBox.Show("Такая запись уже существует!"); }
        }

        private void name_TextChanged(object sender, EventArgs e)
        {
            if (name.Text != "")
            {
                name.BackColor = Color.YellowGreen;
            }

            else name.BackColor = Color.White;

            World();
        }

        private void soder_TextChanged(object sender, EventArgs e)
        {
            if (soder.Text != "")
            {
                soder.BackColor = Color.YellowGreen;
            }

            else soder.BackColor = Color.White;

            World();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fio.Text != "")
            {
                fio.BackColor = Color.YellowGreen;
            }

            else fio.BackColor = Color.White;

            World();
        }
    }
}
