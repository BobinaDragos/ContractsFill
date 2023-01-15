using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Completare_Contracte
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //variabile
        string nrcontract, numeprenume, strada, numar, bloc, scara, etaj, apartament, oras, judet, cnp, serieci, nrci, eliberatde, dataeliberare;
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            workBook.Close();
            Excel.Quit();

            while (Marshal.ReleaseComObject(workBook) != 0) ;
            while (Marshal.ReleaseComObject(Excel) != 0) ;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        int row = 2;
        int col = 1;
        
        static string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\source\ListaMembri.xlsx";
        static Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
        Workbook workBook = Excel.Workbooks.Open(path);
        private void ExcelImport()
        {
            for (int i = 1; i <= workBook.Sheets.Count; i++)
            {
                if (CellValue(row, col) == "") { MessageBox.Show("Nu există date introduse în Excel."); }
                else
                {
                    col++;
                    numeprenume = CellValue(row, col) + " " + CellValue(row, col+1);
                    tnume.Text = numeprenume;
                    col += 5;
                    cnp = CellValue(row, col);
                    tcnp.Text = cnp;
                    col++;
                    serieci = CellValue(row, col);
                    tserie.Text = serieci;
                    col++;
                    nrci = CellValue(row, col);
                    tnrci.Text = nrci;
                    col++;
                    dataeliberare = CellValue(row, col);
                    teliberat.Text = dataeliberare;
                    col++;
                    eliberatde = CellValue(row, col);
                    teliberatde.Text = eliberatde;
                    col++;
                    strada = CellValue(row, col);
                    tstrada.Text = strada;
                    col++;
                    numar = CellValue(row, col);
                    tnumar.Text = numar;
                    col++;
                    bloc = CellValue(row, col);
                    tbloc.Text = bloc;
                    col++;
                    scara = CellValue(row, col);
                    tscara.Text = scara;
                    col++;
                    etaj = CellValue(row, col);
                    tetaj.Text = etaj;
                    col++;
                    apartament = CellValue(row, col);
                    tapartament.Text = apartament;
                    col++;
                    oras = CellValue(row, col);
                    toras.Text = oras;
                    col++;
                    judet = CellValue(row, col);
                    tjudet.Text = judet;
                    col++;
                    nrcontract = CellValue(row, col);
                    tnrcontract.Text = nrcontract;
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelImport();

        }

        private string CellValue (int row,int col)
        {
            string valoareReturnata;
            Worksheet worksheet = workBook.Worksheets[1];
            object cell = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col]).Value;
            if (cell==null){
                return ""; 
            }
            return cell.ToString();
        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object matchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref matchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchDiacritics,
                ref matchControl);
        }
        //Creeate the Doc Method
        private void CreateWordDocument(object filename, object SaveAs)
        {
            if (tnume.Text == "") { MessageBox.Show("Trebuie mai întâi să dai Import."); }
            else
            {
                Word.Application wordApp = new Word.Application();
                object missing = Missing.Value;
                Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    wordApp.Visible = false;

                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing);
                    myWordDoc.Activate();

                    string blocvirg, scaravirg, etajvirg, apartamentvirg;
                    if (tbloc.Text != "")
                    {
                        blocvirg = "Bloc " + tbloc.Text + ", ";
                    }
                    else { blocvirg = tbloc.Text; }

                    if (tscara.Text != "")
                    {
                        scaravirg = "Scara " + tscara.Text + ", ";
                    }
                    else { scaravirg = tscara.Text; }

                    if (tetaj.Text != "")
                    {
                        etajvirg = "Etaj " + tetaj.Text + ", ";
                    }
                    else { etajvirg = tetaj.Text; }

                    if (tapartament.Text != "")
                    {
                        apartamentvirg = "Apartament " + tapartament.Text + ", ";
                    }
                    else { apartamentvirg = tapartament.Text; }

                    //find and replace
                    this.FindAndReplace(wordApp, "<nrcontract>", tnrcontract.Text);
                    this.FindAndReplace(wordApp, "<numeprenume>", tnume.Text);
                    this.FindAndReplace(wordApp, "<strada>", tstrada.Text);
                    this.FindAndReplace(wordApp, "<numar>", tnumar.Text);
                    this.FindAndReplace(wordApp, "<bloc>", blocvirg);
                    this.FindAndReplace(wordApp, "<scara>", scaravirg);
                    this.FindAndReplace(wordApp, "<etaj>", etajvirg);
                    this.FindAndReplace(wordApp, "<apartament>", apartamentvirg);
                    this.FindAndReplace(wordApp, "<oras>", toras.Text);
                    this.FindAndReplace(wordApp, "<judet>", tjudet.Text);
                    this.FindAndReplace(wordApp, "<cnp>", tcnp.Text);
                    this.FindAndReplace(wordApp, "<serieci>", tserie.Text);
                    this.FindAndReplace(wordApp, "<nrci>", tnrci.Text);
                    this.FindAndReplace(wordApp, "<eliberatde>", teliberat.Text);
                    this.FindAndReplace(wordApp, "<dataeliberare>", teliberatde.Text);
                }
                else
                {
                    MessageBox.Show("Nu s-a gasit template-ul!");
                }

                //Save as
                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);

                myWordDoc.Close();
                wordApp.Quit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pathfind = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\source\template.docx";
            string pathsave = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\source\Contracte/" + tnume.Text + ".docx";
            CreateWordDocument(pathfind,pathsave);
            if (tnume.Text == "") { }
            else
            {
                MessageBox.Show("Document creat cu succes!");
                col = 2;
                row++;
                ExcelImport();
            }
        }
        
        

        private void bbulk_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= workBook.Sheets.Count; i++)
            {
                int count = 0;
                while (CellValue(row, col) != "")
                {
                    ExcelImport();
                    string pathfind = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"/source/template.docx";
                    string pathsave = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"/source/Contracte/" + tnume.Text + ".docx";
                    CreateWordDocument(pathfind,pathsave);
                    count++;
                    iter.Text = "Contracte create: " + count;
                    col = 1;
                    row++;
                }
                MessageBox.Show("Contractele au fost completate. Dragoș vă urează o zi minunata! xD");
            }
        }
    }
}