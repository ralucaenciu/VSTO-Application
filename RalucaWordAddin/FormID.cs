using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace RalucaWordAddin
{
    public partial class FormID : Form
    {
       public Angajat angajat;
        DBOperations dboperations;
        public FormID()
        {
            InitializeComponent();
            dboperations = new DBOperations();
            var listAngajati = dboperations.selectAllAngajati();
            dataGridView1.DataSource = listAngajati;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            this.angajat = dboperations.selectAngajat(id);
            dboperations.closeConnection();
            if (angajat != null)
            {
                this.DialogResult = DialogResult.OK; this.Close(); return;
            }
        }
    }
}
