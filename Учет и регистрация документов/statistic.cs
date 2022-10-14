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


using System.Windows.Forms.DataVisualization.Charting;

namespace Учет_и_регистрация_документов
{
    public partial class statistic : Form
    {
        public statistic()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private SQLiteConnection DB;

        public void All()
        {
            SQLiteCommand cmd = DB.CreateCommand();
            cmd.CommandText = "select count(*) from doc";
            cmd.CommandType = CommandType.Text;
            int RowCount = 0;

            RowCount = Convert.ToInt32(cmd.ExecuteScalar());
            label6.Text = RowCount.ToString();      
        }

        public void ispol()
        {
            SQLiteCommand cmd = DB.CreateCommand();
            cmd.CommandText = "select count(*) from doc where Отметка_об_исполнении='ИСПОЛНЕНО'";
            cmd.CommandType = CommandType.Text;
            int RowCount = 0;

            RowCount = Convert.ToInt32(cmd.ExecuteScalar());
            label7.Text = RowCount.ToString();

            double a = RowCount * 100 / Convert.ToDouble(label6.Text);
            label11.Text = "("+a.ToString("0.##") +"%"+")";
        }

        public void ne_ispol()
        {
            SQLiteCommand cmd = DB.CreateCommand();
            cmd.CommandText = "select count(*) from doc where Отметка_об_исполнении='НЕ ИСПОЛНЕНО'";
            cmd.CommandType = CommandType.Text;
            int RowCount = 0;

            RowCount = Convert.ToInt32(cmd.ExecuteScalar());
            label8.Text = RowCount.ToString();

            double a = RowCount * 100 / Convert.ToDouble(label6.Text);
            label12.Text = "(" + a.ToString("0.##") + "%" + ")";
        }

        public void elec()
        {
            SQLiteCommand cmd = DB.CreateCommand();
            cmd.CommandText = "select count(*) from doc where Носитель='ЭЛЕКТРОННЫЙ ДОКУМЕНТ'";
            cmd.CommandType = CommandType.Text;
            int RowCount = 0;

            RowCount = Convert.ToInt32(cmd.ExecuteScalar());
            label9.Text = RowCount.ToString();

            double a = RowCount * 100 / Convert.ToDouble(label6.Text);
            label13.Text = "(" + a.ToString("0.##") + "%" + ")";
        }

        public void bum()
        {
            SQLiteCommand cmd = DB.CreateCommand();
            cmd.CommandText = "select count(*) from doc where Носитель='БУМАЖНЫЙ НОСИТЕЛЬ'";
            cmd.CommandType = CommandType.Text;
            int RowCount = 0;

            RowCount = Convert.ToInt32(cmd.ExecuteScalar());
            label10.Text = RowCount.ToString();


            double a = RowCount * 100 / Convert.ToDouble(label6.Text);
            label14.Text = "(" + a.ToString("0.##") + "%" + ")";
        }

            public void more()
        {
            All();
            ispol();
            ne_ispol();
            elec();
            bum();   
        }


        public void diagramka()
        {         
            chart1.Titles.Add("Диаграмма исполнения документов");
            chart1.Titles[0].Font = new Font("Utopia", 16);

            chart1.Series.Add(new Series("ColumnSeries")
            {
                ChartType = SeriesChartType.Pie
            });

            // Salary series data
            double[] yValues = { Convert.ToDouble(label7.Text), Convert.ToDouble(label8.Text) };
            string[] xValues = { "Исполнено", "Не исполнено", };
            chart1.Series["ColumnSeries"].Points.DataBindXY(xValues, yValues);
        }


        public void diagramka2()
        {
            chart2.Titles.Add("Вид документа");
            chart2.Titles[0].Font = new Font("Utopia", 16);

            chart2.Series.Add(new Series("ColumnSeries")
            {
                ChartType = SeriesChartType.Pie
            });

            // Salary series data
            double[] yValues = { Convert.ToDouble(label9.Text), Convert.ToDouble(label10.Text) };
            string[] xValues = { "Электронный", "Бумажный", };
            chart2.Series["ColumnSeries"].Points.DataBindXY(xValues, yValues);
        }

        private void statistic_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection("Data Source = DocDB.db; Version=3");
            DB.Open();

            more();

            diagramka();
            diagramka2();
        }
    }
}
