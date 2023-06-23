using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RalucaWordAddin
{
    public class DBOperations
    {
        SqlConnection sqlConnection;
       public DBOperations() {
            sqlConnection = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\enciu\\Desktop\\Varianta noua\\RalucaWordAddin\\RalucaWordAddin\\Database1.mdf;Integrated Security=True");
            sqlConnection.Open();

        }

        public void insertAngajat(Angajat angajat) { 

        SqlCommand cmd = new SqlCommand("Insert into dbo.ANGAJATI (NUME,PRENUME,CNP,VARSTA,ADRESA,SEX,DEPARTAMENT,DATANASTERI,POZITIE,SALARIUBRUT,SALARIUNET,CAS,CASS,IV,DP) VALUES (@NUME,@PRENUME,@CNP,@VARSTA,@ADRESA,@SEX,@DEPARTAMENT,@DATANASTERI,@POZITIE,@SALARIUBRUT,@SALARIUNET,@CAS,@CASS,@IV,@DP)", sqlConnection);
           
            cmd.Parameters.AddWithValue("@NUME", angajat.nume);
            cmd.Parameters.AddWithValue("@PRENUME", angajat.prenume);
            cmd.Parameters.AddWithValue("@CNP", angajat.cnp);
            cmd.Parameters.AddWithValue("@VARSTA", angajat.varsta);
            cmd.Parameters.AddWithValue("@ADRESA", angajat.adresa);
            cmd.Parameters.AddWithValue("@SEX",angajat.getSexString());
            cmd.Parameters.AddWithValue("@DEPARTAMENT",angajat.getDepartamentString());
            cmd.Parameters.AddWithValue("@DATANASTERI", angajat.dataNasteri);
            cmd.Parameters.AddWithValue("@POZITIE", angajat.pozitie);
            cmd.Parameters.AddWithValue("@SALARIUBRUT", angajat.salariuBrut);
            cmd.Parameters.AddWithValue("@SALARIUNET", angajat.salariuNet);
            cmd.Parameters.AddWithValue("@CAS", angajat.CAS);
            cmd.Parameters.AddWithValue("@CASS", angajat.CASS);
            cmd.Parameters.AddWithValue("@IV", angajat.IV);
            cmd.Parameters.AddWithValue("@DP", angajat.DP);
            cmd.ExecuteNonQuery();

        }

        public Angajat selectAngajat(int id) {
            string nume = "";
            string prenume = "";
            string cnp = "";
            string adresa = "";
            string departemente = "";
            string pozitie = "";
            double salariuBrut = 0;
            SqlCommand cmd = new SqlCommand("SELECT * FROM ANGAJATI WHERE ID=@ID", sqlConnection);
            cmd.Parameters.AddWithValue("@ID", id);
            SqlDataReader reader = cmd.ExecuteReader();
            while(reader.Read())
            {

                nume = reader["nume"].ToString();
                prenume = reader["prenume"].ToString();
                cnp = reader["cnp"].ToString();
                adresa = reader["adresa"].ToString() ;
                pozitie = reader["pozitie"].ToString();
                salariuBrut = Double.Parse(reader["salariuBrut"].ToString());
            }
            reader.Close();
            Angajat angajat = new Angajat(nume, prenume, cnp, adresa, departemente, pozitie, salariuBrut);
            angajat.id = id;
            return angajat;
        }

        public List<Angajat> selectAllAngajati() {
            var returnList = new List<Angajat>();
            string nume = "";
            string prenume = "";
            string cnp = "";
            string adresa = "";
            string departemente = "";
            string pozitie = "";
            double salariuBrut = 0;
            int id = 0;
            SqlCommand cmd = new SqlCommand("SELECT * FROM ANGAJATI", sqlConnection);
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                id = int.Parse(reader["id"].ToString());
                nume = reader["nume"].ToString();
                prenume = reader["prenume"].ToString();
                cnp = reader["cnp"].ToString();
                adresa = reader["adresa"].ToString();
                pozitie = reader["pozitie"].ToString();
                departemente = reader["departament"].ToString();
                salariuBrut = Double.Parse(reader["salariuBrut"].ToString());
                Angajat angajat = new Angajat(nume, prenume, cnp, adresa, departemente, pozitie, salariuBrut);
                angajat.id = id;
                returnList.Add(angajat);
            }
            reader.Close();
            return returnList;
        }

        public void closeConnection()
        {
            sqlConnection.Close();
        }
    }
}
