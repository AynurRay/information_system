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
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }


        private SQLiteConnection DB;

        private void close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void About_Load(object sender, EventArgs e)
        {
    
        }

    }
}
