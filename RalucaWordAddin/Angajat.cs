using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace RalucaWordAddin
{
    public class Angajat
    {
        public int id { get; set; }
        public string nume { get; set; }
        public string prenume { get; set; }
        public string cnp { get; set; }
        public int varsta { get; set; }
        public string adresa { get; set; }
        public SEX sex { get; set; }
        public DEPARTAMENTE departemente { get; set; }
        public DateTime dataNasteri { get; set; }
        public DateTime dataAngajare { get; set; }
        public string pozitie { get; set; }
        public double salariuBrut { get; set; }
        public double salariuNet { get; set; }
        public double CAS { get; set; }
        public double CASS { get; set; }
        public double IV { get; set; }
        public double DP { get; set; }


        public string getDepartamentString() {
            if (this.departemente == DEPARTAMENTE.IT) {
                return "IT";
            }
            if (this.departemente == DEPARTAMENTE.VANZARI)
            {
                return "VANZARI";
            }
            if (this.departemente == DEPARTAMENTE.FINANTE)
            {
                return "FINANTE";
            }
            if (this.departemente == DEPARTAMENTE.HR)
            {
                return "HR";
            }
            if (this.departemente == DEPARTAMENTE.MARKETING)
            {
                return "MARKETING";
            }
            return null;
        }

        public string getSexString() {
            if (this.sex == SEX.M)
            {

                return "Masculin";

            }

            else {
                return "Feminin";
            }
        }

        public void convertCNP()
        {
            string charSex = cnp.Substring(0, 1);
            if (charSex.Equals("1") || charSex.Equals("5"))
            {
                this.sex = SEX.M;
            }
            else if (charSex.Equals("2") || charSex.Equals("6"))
            {
                this.sex = SEX.F;
            }
            string dataNasteri = cnp.Substring(1, 6);
            string prefixan = "";
            if (charSex.Equals("1") || charSex.Equals("2"))
            {
                prefixan = "19";
            }
            else { prefixan = "20"; }
            dataNasteri = prefixan + dataNasteri;
            this.dataNasteri =DateTime.ParseExact(dataNasteri,
                                  "yyyyMMdd",
                                   CultureInfo.InvariantCulture);
            this.varsta = DateTime.Now.Year - this.dataNasteri.Year;

        }

        public void calculSalarial()
        {
           // double salariuBrutIncrease = 50; 
           // double DPDecrease = 15; 
            this.CAS = ((double)25 / 100) * this.salariuBrut;
            this.CASS = ((double)10 / 100) * this.salariuBrut;
            this.IV = ((double)10 / 100) * this.salariuBrut;
            this.DP = 0;
            this.salariuNet = this.salariuBrut - CAS - CASS - IV - DP;
        }

        public Angajat(string nume, string prenume, string cnp, string adresa, string departemente, string pozitie, double salariuBrut)
        {
            this.nume = nume;
            this.prenume = prenume;
            this.cnp = cnp;
            this.adresa = adresa;
            switch (departemente) {
                case "IT":
                    this.departemente = DEPARTAMENTE.IT;
                    break;
                case "VANZARI"
                :this.departemente = DEPARTAMENTE.VANZARI;
                    break;
                case "FINANTE":
                    this.departemente = DEPARTAMENTE.FINANTE;
                    break;
                case "HR":
                    this.departemente = DEPARTAMENTE.HR;
                    break;
                case "MARKETING":
                    this.departemente = DEPARTAMENTE.MARKETING;
                    break;
                default:
                    this.departemente=DEPARTAMENTE.IT;
                    break;
            }
            this.pozitie = pozitie;
            this.salariuBrut = salariuBrut;
            this.convertCNP();
            this.calculSalarial();
        }
    }

    public enum SEX
    {
        M,F
    }

    public enum DEPARTAMENTE
    {
        IT,VANZARI,FINANTE,HR,MARKETING
    }
}
