using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RalucaWordAddin
{
    public partial class Form1 : Form
    {
        public Angajat angajat;
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double salariu = 0;
            if (textBoxNume.Text.Length < 2) {
                MessageBox.Show("Numele trebuie sa fie mai mare de 2 caractere");
                return;
            }
            if (textBoxPrenume.Text.Length < 2)
            {
                MessageBox.Show("Preumele trebuie sa fie mai mare de 2 caractere");
                return;
            }
            if (textBoxCNP.Text.Length != 13)
            {
                MessageBox.Show("CNP-ul nu este valid");
                return;
            }
            if (textBoxAdresa.Text.Length < 2)
            {
                MessageBox.Show("Adresa trebuie sa fie mai mare de 3 caractere");
                return;
            }
            if (textBoxPozitie.Text.Length < 2)
            {
                MessageBox.Show("Pozitia trebuie sa fie mai mare de 3 caractere");
                return;
            }
            if (textBoxSalariuBrut.Text.Length > 2)
            {
                string salBrut = textBoxSalariuBrut.Text;
                salBrut = salBrut.Trim();
                try
                {
                    salariu = Double.Parse(salBrut);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Salariul trebuie sa aiba numai numere");
                    return;
                }
            }
            else {
                MessageBox.Show("Salariul brut trebuie sa aiba MINIM 2 cifre");
                return;
            }
            this.angajat=new Angajat(textBoxNume.Text,textBoxPrenume.Text,textBoxCNP.Text,textBoxAdresa.Text,comboBoxDepartamente.SelectedItem.ToString(),textBoxPozitie.Text,salariu);
            this.DialogResult = DialogResult.OK;
            this.Close();

        }

        private void comboBoxDepartamente_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void textBoxSalariuBrut_TextChanged(object sender, EventArgs e)
        {
        }
    }
}
