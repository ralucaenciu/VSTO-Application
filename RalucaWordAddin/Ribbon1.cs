using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;

namespace RalucaWordAddin
{
    public partial class Ribbon1
    {
        Microsoft.Office.Interop.Word.Document document;
        Microsoft.Office.Interop.Word.Paragraph paragraphDate;
        DBOperations dBOperations;
        Angajat angajat;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            dBOperations = new DBOperations();
            this.angajat = new Angajat("Raluca", "Maria", "5000519450037", "asdadd", "IT", "SENIOR", 999999);
        }

        private void btnAdaugaAngajat_Click(object sender, RibbonControlEventArgs e)
        {
            document = Globals.ThisAddIn.Application.ActiveDocument;
             Form1 form = new Form1();
             form.ShowDialog();
              if (form.DialogResult == DialogResult.OK)
              {
                  this.angajat = form.angajat;
            incarcarePagina();
             }
        }

        private void incarcarePagina() {
            document = Globals.ThisAddIn.Application.ActiveDocument;
            foreach(Paragraph paragraph in document.Paragraphs)
            {
                paragraph.Range.Delete();
            }
            var paraTitle = document.Application.ActiveDocument.Paragraphs.Add();
            paraTitle.Range.Text = "Document pentru veridicitatea datelor angajatului " + angajat.nume + " " + angajat.prenume;
            paraTitle.Range.Font.Size = 30;
            paraTitle.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paraTitle.Range.InsertParagraphAfter();
            var paraDisclamier = document.Application.ActiveDocument.Paragraphs.Add();
            paraDisclamier.Range.Text = "Acest document este confidential, va rugam sa nu impartasiti aceste informatii cu nimeni altcineva din companie si din afara ei. Va rugam de asememena " +
              "va rugam sa verificati corectitudinea datelor trimise, acestea odata intrate in baza de date sunt mai dificil de modificat. Pentru confirmarea celor citite, va rugam sa trimiteti email la hrDigital@dg.com cu documentul semnat. " +
              "Va multumim si va uram o zi buna!";
            paraDisclamier.Range.Font.Size = 12;
            paraDisclamier.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paraDisclamier.Range.InsertParagraphAfter();
            asignareDateWord("Nume angajat", angajat.nume);
            asignareDateWord("Prenume angajat", angajat.prenume);
            asignareDateWord("Adresa angajat", angajat.adresa);
            asignareDateWord("C.N.P", angajat.cnp);
            asignareDateWord("Data nastere", angajat.dataNasteri.ToShortDateString());
            asignareDateWord("Varsta", angajat.varsta.ToString());
            asignareDateWord("Pozitie", angajat.pozitie);
            asignareDateWord("Departament", angajat.getDepartamentString());
            asignareDateWord("CAS", angajat.CAS.ToString());
            asignareDateWord("CASS", angajat.CASS.ToString());
            asignareDateWord("IV", angajat.IV.ToString());
            asignareDateWord("Salariu Brut", angajat.salariuBrut.ToString());
            asignareDateWord("Salariu Net", angajat.salariuNet.ToString());
            asignareDateWord("Sex", angajat.getSexString());
            var rangeFinal = document.Application.ActiveDocument.Paragraphs.Add().Range;
            rangeFinal.Text = "Data: " + "\n" + DateTime.Now.ToShortDateString() + "\nSemnatura:" + "\n...............";
            rangeFinal.Paragraphs.Format.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
        
            document.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields.Add(
        document.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range,
        Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
        }

        public void asignareDateWord(string paragrafNume, string paragrafVal) {
            paragraphDate = document.Application.ActiveDocument.Paragraphs.Add();
            paragraphDate.Range.Font.Name = "Times New Roman";
            paragraphDate.Range.Font.Size = 12;
            paragraphDate.Range.Text = paragrafNume + ": " + paragrafVal;
            object objBoldStart = paragraphDate.Range.Start;
            object objBoldEnd = paragraphDate.Range.Start + paragraphDate.Range.Text.IndexOf(":");
            var objBoldRange=document.Range(ref objBoldStart, ref objBoldEnd);
            objBoldRange.Bold = 1;
            object objItalicStart = paragraphDate.Range.Start + paragraphDate.Range.Text.IndexOf(":");
            object objItalicEnd = paragraphDate.Range.End;
            var objItalicRange=document.Range(ref objItalicStart, ref objItalicEnd);
            objItalicRange.Italic = 1;
            paragraphDate.Range.InsertParagraphAfter();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string nume="";
            string prenume="";
            string cnp = "";
            string adresa = "";
            string departemente = "";
            string pozitie = "";
            double salariuBrut=0;
            foreach (Microsoft.Office.Interop.Word.Paragraph aPar in document.Application.ActiveDocument.Paragraphs) {
                int indexDoublePoints = 0;
                if ((indexDoublePoints=aPar.Range.Text.IndexOf(": ")) != -1) {
                    string value= aPar.Range.Text.Substring(indexDoublePoints+2);
                    value=value.Trim();
                    if (aPar.Range.Text.StartsWith("Nume angajat")) {
                        nume = value;
                        continue;
                    }
                    if (aPar.Range.Text.StartsWith("Prenume angajat"))
                    {
                        prenume = value;
                        continue;
                    }
                    if (aPar.Range.Text.StartsWith("Adresa angajat"))
                    {
                        adresa = value;
                        continue;
                    }
                    if (aPar.Range.Text.StartsWith("C.N.P"))
                    {
                        cnp = value;
                        continue;   
                    }
                    if (aPar.Range.Text.StartsWith("Pozitie"))
                    {
                       pozitie = value;
                        continue;
                    }
                    if (aPar.Range.Text.StartsWith("Salariu Brut"))
                    {
                        salariuBrut = Double.Parse(value);
                        continue;
                    }
                    if (aPar.Range.Text.StartsWith("Departament")) {
                        departemente = value;
                        continue;
                    }
                }
            }
            Angajat nouAngajat = new Angajat(nume, prenume, cnp, adresa, departemente, pozitie, salariuBrut);
            this.angajat = nouAngajat;
            dBOperations.insertAngajat(this.angajat);
            MessageBox.Show("Inserare angajat cu succes");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            FormID formID = new FormID();
            formID.ShowDialog();
            if (formID.DialogResult == DialogResult.OK) {
                this.angajat = formID.angajat;
                incarcarePagina();
                return;
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            List<Angajat> angajati = dBOperations.selectAllAngajati();
            string projectDirectory = "C:\\AddIn1";
            var excelFilePath = projectDirectory + "\\excel.xlsx";
            var excelApp= new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Open(excelFilePath);
            var worksheet = workbook.Sheets[1];
            worksheet.Cells.ClearContents();
            if (worksheet.ChartObjects().Count > 0)
            {
                var chart1 = worksheet.ChartObjects(1);
                if (chart1 != null)
                {
                    chart1.Delete();
                }
            }
            var dataRange = worksheet.Range["B2"];
            dataRange.Value2 = 1;
          
            (worksheet.Cells[1, 2] as Microsoft.Office.Interop.Excel.Range).Value2 = "Id";
            (worksheet.Cells[1, 3] as Microsoft.Office.Interop.Excel.Range).Value2 = "Nume";
            (worksheet.Cells[1, 4] as Microsoft.Office.Interop.Excel.Range).Value2 = "Prenume";
            (worksheet.Cells[1, 5] as Microsoft.Office.Interop.Excel.Range).Value2 = "CNP";
            (worksheet.Cells[1, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = "Varsta";
            (worksheet.Cells[1, 7] as Microsoft.Office.Interop.Excel.Range).Value2 = "Strada";
            (worksheet.Cells[1, 8] as Microsoft.Office.Interop.Excel.Range).Value2 = "Sex";
            (worksheet.Cells[1, 9] as Microsoft.Office.Interop.Excel.Range).Value2 = "Departament";
            (worksheet.Cells[1, 10] as Microsoft.Office.Interop.Excel.Range).Value2 = "Data nastere";
            (worksheet.Cells[1, 11] as Microsoft.Office.Interop.Excel.Range).Value2 = "Pozitie";
            (worksheet.Cells[1, 12] as Microsoft.Office.Interop.Excel.Range).Value2 = "Salariu Brut";
            (worksheet.Cells[1, 13] as Microsoft.Office.Interop.Excel.Range).Value2 = "Salariu Net";
            (worksheet.Cells[1, 14] as Microsoft.Office.Interop.Excel.Range).Value2 = "CAS";
            (worksheet.Cells[1, 15] as Microsoft.Office.Interop.Excel.Range).Value2 = "CASS";
            (worksheet.Cells[1, 16] as Microsoft.Office.Interop.Excel.Range).Value2 = "IV";
       
            for (int i = 0; i < angajati.Count; i++)
            {

                (worksheet.Cells(i + 2, 2) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].id;
                (worksheet.Cells(i + 2, 3) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].nume;
                (worksheet.Cells(i + 2, 4) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].prenume;
                (worksheet.Cells(i + 2, 5) as Microsoft.Office.Interop.Excel.Range).Value2 = "\'"+ angajati[i].cnp.ToString();
                (worksheet.Cells(i + 2, 6) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].varsta;
                (worksheet.Cells(i + 2, 7) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].adresa;
                (worksheet.Cells(i + 2, 8) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].sex.ToString();
                (worksheet.Cells(i + 2, 9) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].departemente.ToString();
                (worksheet.Cells(i + 2, 10) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].dataNasteri.ToShortDateString();
                (worksheet.Cells(i + 2, 11) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].pozitie;
                (worksheet.Cells(i + 2, 12) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].salariuBrut;
                (worksheet.Cells(i + 2, 13) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].salariuNet;
                (worksheet.Cells(i + 2, 14) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].CAS;
                (worksheet.Cells(i + 2, 15) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].CASS;
                (worksheet.Cells(i + 2, 16) as Microsoft.Office.Interop.Excel.Range).Value2 = angajati[i].IV;
            }

            var chartRange = worksheet.Range["L1:M" + (angajati.Count + 1).ToString()];
            var charts = worksheet.ChartObjects() as Microsoft.Office.Interop.Excel.ChartObjects;
            var chartObject = charts.Add(800, 10, 300, 300) as Microsoft.Office.Interop.Excel.ChartObject;
            var chart = chartObject.Chart;
            chart.SetSourceData(chartRange);
            chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered;


            Microsoft.Office.Interop.Excel.Range rangeMin = worksheet.Range["M:M"];
            double minValue = (double)worksheet.Application.WorksheetFunction.Min(rangeMin);
            (worksheet.Cells[14, 4] as Microsoft.Office.Interop.Excel.Range).Value2 = "Salariu net minim";
            worksheet.Cells[15, 4] = minValue;

            Microsoft.Office.Interop.Excel.Range rangeMax = worksheet.Range["M:M"];
            double maxValue = (double)worksheet.Application.WorksheetFunction.Max(rangeMax);
            (worksheet.Cells[14, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = "Salariu net maxim";
            worksheet.Cells[15, 6] = maxValue;


            Microsoft.Office.Interop.Excel.Range avgSal = worksheet.Range["L:L"];
            double avgSalBrut = (double)worksheet.Application.WorksheetFunction.Average(avgSal);
            (worksheet.Cells[14, 8] as Microsoft.Office.Interop.Excel.Range).Value2 = "Medie salariu brut";
            worksheet.Cells[15, 8] = avgSalBrut;


            workbook.Close(SaveChanges: true, excelFilePath);
            excelApp.Quit();
            var process = System.Diagnostics.Process.Start(excelFilePath);
            process.Close();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string pdfFilePath = "C:\\AddIn1\\document.pdf";
            Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            document.ExportAsFixedFormat(pdfFilePath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
        }
    }
}
