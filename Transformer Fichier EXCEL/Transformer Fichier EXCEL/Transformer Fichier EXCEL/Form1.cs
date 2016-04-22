using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

using System.Runtime.InteropServices;
using System.Collections;
using INI;

using System.IO;
using System.Xml;
using System.Threading;
using System.Diagnostics;

namespace TransformEXCEL
{
    public partial class Form1 : Form   
    {
       bool flagsimplypastout=true;
        string stylexml = null;
        string fichierprepare = null;
        string pathnotapme = null;
        string divitylerfinal = null;
        string pathstylerfinal = null;
        /// <summary>
        /// //////////////////////////////
        /// </summary>
        string prefaceNP = null;
        string fileAstyler = null;

        string timfusion = "init";
        string timleger = "init";

        string timdiviser = "0";
        string timdiviserHistoS = "0";
        string timdiviserAnnuel = "0";
        string timdiviserSynthese = "0";

        string timtotal = "init";

        public Form1()
        {
            InitializeComponent();
            string file1 = "D:\\ptw\\Annuel.ptw";
            string file2 = "D:\\ptw\\Admin.ptw";
            string file3 = "D:\\ptw\\Histo.ptw";
            string file4 = "D:\\ptw\\Eval.ptw";
            string file5 = "D:\\ptw\\Decis.ptw";
            string file6 = "D:\\ptw\\Tres.ptw";
            string file7 = "D:\\ptw\\Histo-s.ptw";

            if (File.Exists(file1))
            {
                FileInfo fi = new FileInfo(file1);
                checkBox12.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox12.Checked = true;
            }
            if (File.Exists(file2))
            {
                FileInfo fi = new FileInfo(file2);
                checkBox13.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox13.Checked = true;
            }
            if (File.Exists(file3))
            {
                FileInfo fi = new FileInfo(file3);
                checkBox14.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox14.Checked = true;
            }
            if (File.Exists(file4))
            {
                FileInfo fi = new FileInfo(file4);
                checkBox15.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox15.Checked = true;
            }
            if (File.Exists(file5))
            {
                FileInfo fi = new FileInfo(file5);
                checkBox16.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox16.Checked = true;
            }
            if (File.Exists(file6))
            {
                FileInfo fi = new FileInfo(file6);
                checkBox17.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox17.Checked = true;
            }
            if (File.Exists(file7))
            {
                FileInfo fi = new FileInfo(file7);
                checkBox18.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox18.Checked = true;
            }
            //Initialize path info pour le programme
            ////Renvoyer en *.ini ?
            string filePath = "D:\\ptw\\configfile\\pathinfo.ini";
            IniFile iniFile = new IniFile(filePath);

            string pathsource = null;
            string pathxml = null;
            string pathdestinationdivi = null;
            string pathdestinationstyle = null;
            string pathdestinationfusion = null;
            string pathSourceFusion = null;
            string sourcestyle = null;
            string sourcedivi = null;
            string styledivi = null;
            string sourceprefaceNP = null;
            string pathprefaceNP = null;
            string stylefusion = null;
            string styletest = null;
            string sourcestyletest = null;
            string sourceindex = null;
            string pathdestinationindex = null;
            string pathdelockedstatus = null;
            pathsource = iniFile.ReadInivalue("dossier", "pathsource");
            pathxml = iniFile.ReadInivalue("dossier", "pathxml");
            pathdestinationdivi = iniFile.ReadInivalue("dossier", "pathdestinationdivi");
            pathdestinationstyle = iniFile.ReadInivalue("dossier", "pathdestinationstyle");
            pathdestinationfusion = iniFile.ReadInivalue("dossier", "pathdestinationfusion");
            pathSourceFusion = iniFile.ReadInivalue("dossier", "pathSourceFusion");
            sourcestyle = iniFile.ReadInivalue("dossier", "sourcestyle");
            sourcedivi = iniFile.ReadInivalue("dossier", "sourcedivi");
            styledivi = iniFile.ReadInivalue("dossier", "styledivi");
            sourceprefaceNP = iniFile.ReadInivalue("dossier", "sourceprefaceNP");
            pathprefaceNP = iniFile.ReadInivalue("dossier", "pathprefaceNP");
            stylefusion = iniFile.ReadInivalue("dossier", "stylefusion");
            styletest = iniFile.ReadInivalue("dossier", "styletest");
            sourcestyletest = iniFile.ReadInivalue("dossier", "sourcestyletest");
            sourceindex = iniFile.ReadInivalue("dossier", "sourceindex");
            pathdelockedstatus = iniFile.ReadInivalue("dossier", "pathdelockedstatus");
            pathdestinationindex = iniFile.ReadInivalue("dossier", "pathdestinationindex");
            textBox1.Text = pathsource;
            textBox2.Text = pathxml;
            textBox3.Text = pathdestinationdivi;
            textBox6.Text = pathdestinationstyle;
            textBox5.Text = pathdestinationfusion;
            textBox4.Text = pathSourceFusion;
            textBox7.Text = sourcestyle;
            textBox8.Text = styletest;
            textBox9.Text = sourcedivi;
            textBox19.Text = sourcedivi;
            textBox10.Text = styledivi;
            textBox11.Text = sourceprefaceNP;
            textBox12.Text = pathprefaceNP;
            textBox13.Text = stylefusion;
            textBox14.Text = sourcestyletest;
            textBox15.Text = sourceindex;
            textBox16.Text = pathdestinationindex;
            textBox17.Text = pathdelockedstatus;
            textBox18.Text = sourceprefaceNP;

        }

        #region création de PrefaceNP.xls et Preface.xls
        //
        //// Supprimer l'onglet Typologie IFRS
        //
        private void supprimerTypologie_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(textBox11.Text, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            
            //Ne pas afficher les alertes !! ne pas mettre à False avant de s'assurer du bon fonctionnement !!!
            xlApp.DisplayAlerts = false;

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet sheetTypologie = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Typologie IFRS");
            sheetTypologie.Delete();


            xlWorkBook.SaveAs(prefaceNP, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue,
            misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange,
            misValue, misValue, misValue, misValue, misValue);

            //sauvegarder sans quitter l'insatnce Excel
            //xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //sauvegarder et quitter
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Fonction Remplacement dans Histo.ptw des formules qui génèreront des erreurs après suppressions des lignes IFRS
        //
        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            //string findxo = "+Historique!C**Hist.Preface!D$14";
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Range range = xlWorkSheet.UsedRange;
            // In the following cases Value2 returns different types                
            // 1. the range variable points to a single cell                
            // Value2 returns a object                
            // 2. the range variable points to many cells                
            // Value2 returns object[,] 
            object[,] values = (object[,])range.Formula;
            
            //---------->  \+Historique\!C\d{2,4}\*Hist.Preface\!D\$14 <-----------------
            //Alex: questions........................................................
            xlApp.DisplayAlerts = false;
            range.Cells.Replace("SI(Hist.Preface!D$14=0", "SI(0=0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("SI(Hist.Preface!F$14=0", "SI(0=0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("SI(Hist.Preface!H$14=0", "SI(0=0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("Historique!C?????~*Hist.Preface!D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C????~*Hist.Preface!D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C???~*Hist.Preface!D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C??~*Hist.Preface!D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!D?????~*Hist.Preface!F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D????~*Hist.Preface!F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D???~*Hist.Preface!F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D??~*Hist.Preface!F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!E?????~*Hist.Preface!H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E????~*Hist.Preface!H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E???~*Hist.Preface!H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E??~*Hist.Preface!H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Hist.Calculs remplacement CDE14 
        //
        private void HistoCalculs()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Range range = xlWorkSheet.UsedRange;
            //pour convertir dans les formules cellule relative à absolue (EX: D14 ----  $D$14)
            //Alex: questions

            range.Cells.Replace("Hist.Preface!D14", "Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!$D14", "Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!D$14", "Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!F14", "Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!$F14", "Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!F$14", "Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!H14", "Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!$H14", "Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!H$14", "Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            //remplacer tous les éléments de formules multipliés par $D$14, $F$14, $H$14 par "0"
            range.Cells.Replace("Historique!C?????~*Hist.Preface!$D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C????~*Hist.Preface!$D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C???~*Hist.Preface!$D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C??~*Hist.Preface!$D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!D?????~*Hist.Preface!$F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D????~*Hist.Preface!$F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D???~*Hist.Preface!$F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D??~*Hist.Preface!$F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!E?????~*Hist.Preface!$H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E????~*Hist.Preface!$H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E???~*Hist.Preface!$H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E??~*Hist.Preface!$H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);



            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// HistoMettreZero pour col 8000
        //
        private void HistoMettreZero_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;
            int col = 0;
            string headercol = "";

            //Nouvelle fonction pour trouver le numéro de col "8000"
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col = cf.FindCodedColumn("8000", range);
            headercol = cf.FindCodedColumnHeader("85000", range);
            range.Cells.Replace("Historique!C?????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C???~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C??~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!D?????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D???~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D??~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!E?????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E????~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E???~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E??~*Hist.Refer!$E$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!C?????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C???~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C??~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!D?????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D???~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D??~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!E?????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E????~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E???~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E??~*Hist.Refer!$H$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("Historique!C?????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C???~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!C??~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!D?????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D???~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!D??~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("Historique!E?????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E????~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E???~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Historique!E??~*Hist.Refer!$K$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //>>>>>>>>>>>>>L'ancienne fonction à supprimer :
            //int rCnt = 0;
            //int cCnt = 0;
            //rCnt = range.Rows.Count;
            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "8000"))
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}
            

            //MessageBox.Show(col.ToString());
            //int col = 8;
            int time1 = System.Environment.TickCount;
            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "-1"))
                {
                    //MessageBox.Show(row.ToString());
                    Excel.Range rangeDelxC = xlWorkSheet.Cells[row, 3] as Excel.Range;
                    rangeDelxC.Value2 = 0;
                    Excel.Range rangeDelxD = xlWorkSheet.Cells[row, 4] as Excel.Range;
                    rangeDelxD.Value2 = 0;
                    Excel.Range rangeDelxE = xlWorkSheet.Cells[row, 5] as Excel.Range;
                    rangeDelxE.Value2 = 0;
                    //rangeDelxE.set_Value(misValue, 0);
                }
            }
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + "seconds used");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Historique Remplacement de Hist.Preface $D$14 par 0
        //
        private void HistoRempl_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Formula;

            xlApp.DisplayAlerts = false;

            range.Cells.Replace("C???~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("C????~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D???~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D????~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E???~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E????~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("C???~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("C????~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D???~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D????~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E???~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E????~*Hist.Preface!$F$14", "0*Hist.Preface!$F$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("C???~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("C????~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D???~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("D????~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E???~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("E????~*Hist.Preface!$H$14", "0*Hist.Preface!$H$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            range.Cells.Replace("K????~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("K?????~*Hist.Preface!$D$14", "0*Hist.Preface!$D$14", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

           
            
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //MessageBox.Show("jobs done!");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Historique Au Av Aw Histo3 // a voir l'ancien fichier histo.ptw 
        //
        private void HistoAuAvAw_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Formula;

            range.Cells.Replace("0~*AU??", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("0~*AV??", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("0~*AW??", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("AU??~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AV??~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AW??~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("AU???~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AV???~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AW???~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range.Cells.Replace("AU????~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AV????~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("AW????~*0", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //MessageBox.Show("jobs done!");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Colonne CE IFRS mettre à "0" pour les non null 72000 Histo4.xls
        //
        private void colCE_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");

            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            rCnt = range.Rows.Count;
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "72000"))
                {
                    col = cCnt;
                    break;
                }
            }

            int time1 = System.Environment.TickCount;

            for (int row = 1; row <= values.GetUpperBound(0)-2; row++)//-1-2
            {
                string value = Convert.ToString(values[row, col]);
                //string val = rangeDelx.Value2.ToString();
                if (value != "")
                {
                    //MessageBox.Show(value);
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    rangeDelx.set_Value(misValue, 0);
                    if (value == "72000")
                    {
                        Excel.Range ce72000 = xlWorkSheet.Cells[row, col] as Excel.Range;
                        ce72000.set_Value(misValue, "72000");
                    }
                }
            }

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + "seconds used");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }
        // 
        //// suppression des REF! Histo5.xls
        //
        private void supprimerREF_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;
            int time1 = System.Environment.TickCount;

            //////////////////////////////////////////////////////////////////////////
            xlApp.DisplayAlerts = false;

            range.Cells.Replace("Hist.Preface!$D$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!$F$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!$H$14", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Calculs!$B$4", "0,001", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //Alex: fixed

            range.Cells.Replace("Hist.Preface!F13", "1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("Hist.Preface!H13", "1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            ////////////////////////
            range.Cells.Replace("$AJ$2=10", "0=10", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("$AJ$2=2", "0=2", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            ////////////////////////
            range.Cells.Replace("'*Param Sav'!$C$280=0", "0=0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range.Cells.Replace("'*Param Sav'!$C$183=1", "0=1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            ////////////////////////
            //range.Cells.Replace("$EL$1", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$2", "1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$3", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$4", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$5", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false); 
            //range.Cells.Replace("$EL$7", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$10", "1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //range.Cells.Replace("$EL$11", "0", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            ////////////////////////
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }
        // 
        //// Insérer les colonnes Correctifs // routine hist.Refer
        //
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
         //   prefaceNP = @"C:\Users\Ordimega\Downloads\prefaceNP.xlsx";
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //hist
            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets[1] as Excel.Worksheet;
            Excel.Range range = xlWorkSheet.UsedRange;
       
            Excel.Range rangex1 = xlWorkSheet.Cells[1, 4] as Excel.Range;

            Excel.Range rangex2 = xlWorkSheet.Cells[1, 5] as Excel.Range;

            Excel.Range rangex3 = xlWorkSheet.Cells[1, 6] as Excel.Range;

            //hist.refer
            Excel.Worksheet xlWorkSheet2 = xlWorkBook.Worksheets["Hist.Refer"] as Excel.Worksheet;
            int RowC = xlWorkSheet2.UsedRange.Rows.Count - 1;


            CodeFinder cfRef;
            cfRef = new CodeFinder(xlWorkBook, xlWorkSheet2);
            string colC = cfRef.FindCodedColumnHeader("3000", xlWorkSheet2.UsedRange);
            string colD = cfRef.FindCodedColumnHeader("3000-1000", xlWorkSheet2.UsedRange);
            string colE = cfRef.FindCodedColumnHeader("3000-2000", xlWorkSheet2.UsedRange);
            string colF = cfRef.FindCodedColumnHeader("4000", xlWorkSheet2.UsedRange);
            string colG = cfRef.FindCodedColumnHeader("4000-1000", xlWorkSheet2.UsedRange);
            string colH = cfRef.FindCodedColumnHeader("4000-2000", xlWorkSheet2.UsedRange);
            string colI = cfRef.FindCodedColumnHeader("5000", xlWorkSheet2.UsedRange);
            string colJ = cfRef.FindCodedColumnHeader("5000-1000", xlWorkSheet2.UsedRange);
            string colK = cfRef.FindCodedColumnHeader("5000-2000", xlWorkSheet2.UsedRange);


            Excel.Range rangeCc = xlWorkSheet2.UsedRange.get_Range(colC + "1", xlWorkSheet2.Cells[RowC, 3]) as Excel.Range;
            Excel.Range rangeCc2 = xlWorkSheet2.UsedRange.get_Range(colD + "1", xlWorkSheet2.Cells[RowC, 4]) as Excel.Range;
            Excel.Range rangeCc3 = xlWorkSheet2.UsedRange.get_Range(colE + "1", xlWorkSheet2.Cells[RowC, 5]) as Excel.Range;
            Excel.Range rangeDc = xlWorkSheet2.UsedRange.get_Range(colF + "1", xlWorkSheet2.Cells[RowC, 6]) as Excel.Range;
            Excel.Range rangeDc2 = xlWorkSheet2.UsedRange.get_Range(colG + "1", xlWorkSheet2.Cells[RowC, 7]) as Excel.Range;
            Excel.Range rangeDc3 = xlWorkSheet2.UsedRange.get_Range(colH + "1", xlWorkSheet2.Cells[RowC, 8]) as Excel.Range;
            Excel.Range rangeEc = xlWorkSheet2.UsedRange.get_Range(colI + "1", xlWorkSheet2.Cells[RowC, 9]) as Excel.Range;
            Excel.Range rangeEc2 = xlWorkSheet2.UsedRange.get_Range(colJ + "1", xlWorkSheet2.Cells[RowC, 10]) as Excel.Range;
            Excel.Range rangeEc3 = xlWorkSheet2.UsedRange.get_Range(colK + "1", xlWorkSheet2.Cells[RowC, 11]) as Excel.Range;

            
           
            //hist.preface
            Excel.Worksheet xlWorkSheetpre = xlWorkBook.Worksheets["Hist.Preface"] as Excel.Worksheet;


            CodeFinder cfcol;
            cfcol = new CodeFinder(xlWorkBook, xlWorkSheetpre);


            //hist.calcule
            Excel.Worksheet WorkSheetCalculs = xlWorkBook.Worksheets["Hist.Calculs"] as Excel.Worksheet;
            Excel.Range rangexh1 = WorkSheetCalculs.Cells[1, 1] as Excel.Range;
           
            //l'ordre de déclaration à respecter
           

            //rangex1c.EntireColumn.Copy(misValue);
            //rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangex1c.EntireColumn.Copy(misValue);
            //rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangex1c.EntireColumn.Copy(misValue);
            //rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //
            string colcorrectif1 = "";
            string colcorrectif2 = "";
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);

           


            //l'ordre de déclaration à respecter
            //hist 1st year insert
            string insert1 = cf.FindCodedColumnHeader("4000", range);
            rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            string insert12 = cf.FindCodedColumnHeader("4000", range);
            rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //hist.preface insert 1st
            string colDpre = cfcol.FindCodedColumnHeader("4000", xlWorkSheetpre.UsedRange);
            string colEpre = cfcol.FindCodedColumnHeader("5000", xlWorkSheetpre.UsedRange);
            Excel.Range rangeinsert1 = xlWorkSheetpre.UsedRange.get_Range(colDpre + "1", colEpre + "1") as Excel.Range;
            rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //hist.calcul insert 1st
            rangexh1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangexh1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //hist 1st cut
            xlWorkSheet.get_Range("C1", "C1").EntireColumn.Cut(xlWorkSheet.get_Range("E1", "E1").EntireColumn);
            //hist.refer 1st cut
            rangeCc.Cut(rangeCc3);
            //hist.preface 1st cut

            //hist 2ed insert
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            string insert2 = cf.FindCodedColumnHeader("5000", range);
            rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            string insert22 = cf.FindCodedColumnHeader("5000", range);
            rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //hist.preface 2ed insert
            string colHpre = cfcol.FindCodedColumnHeader("6000", xlWorkSheetpre.UsedRange);
            string colIpre = cfcol.FindCodedColumnHeader("7000", xlWorkSheetpre.UsedRange);
            Excel.Range rangeinsert2 = xlWorkSheetpre.UsedRange.get_Range(colHpre + "1", colIpre + "1") as Excel.Range;
            rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //hist.calculs 2ed insert
            Excel.Range rangexh2 = WorkSheetCalculs.Cells[1, 5] as Excel.Range;
            rangexh2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangexh2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //hist 2ed cut
            xlWorkSheet.get_Range("F1", "F1").EntireColumn.Cut(xlWorkSheet.get_Range("H1", "H1").EntireColumn);
            //hist.refer 2ed cut
            rangeDc.Cut(rangeDc3);
            //hist.preface 2ed cut





           

            //hist 3ed insert
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            string insert3 = cf.FindCodedColumnHeader("6000", range);
            rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            string insert32 = cf.FindCodedColumnHeader("6000", range);
            rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //hist.preface 3ed insert
            string colL = cfcol.FindCodedColumnHeader("8000", xlWorkSheetpre.UsedRange);
            string colM = cfcol.FindCodedColumnHeader("9000", xlWorkSheetpre.UsedRange);
            Excel.Range rangeinsert3 = xlWorkSheetpre.UsedRange.get_Range(colL + "1", colM + "1") as Excel.Range;
            rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //hist.calculs 3ed insert
            Excel.Range rangexh3 = WorkSheetCalculs.Cells[1, 9] as Excel.Range;
            rangexh3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangexh3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //hist 3ed cut
            xlWorkSheet.get_Range("I1", "I1").EntireColumn.Cut(xlWorkSheet.get_Range("K1", "K1").EntireColumn);
            //hist.refer 3ed cut
            rangeEc.Cut(rangeEc3);
            //hist.preface 3ed cut
            colcorrectif1 = cf.FindCodedColumnHeader("82000-1000", range);
            colcorrectif2 = cf.FindCodedColumnHeader("82000-2000", range);
            xlWorkSheet.UsedRange.get_Range(colcorrectif1 + "1", colcorrectif1 + "1").EntireColumn.Replace("Hist.Refer!EX", "Hist.Refer!" + colcorrectif1, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //hist 1st copy
            xlWorkSheet.get_Range("E1", "E1").EntireColumn.Copy(xlWorkSheet.get_Range("C1", "C1").EntireColumn);
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            colcorrectif1 = cf.FindCodedColumnHeader("82000-1000", range);
            colcorrectif2 = cf.FindCodedColumnHeader("82000-2000", range);
            Excel.Range rangex1c = xlWorkSheet.UsedRange.get_Range(colcorrectif1 + "1", colcorrectif2 + "1") as Excel.Range;
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range(insert1 + "1", insert12 + "1").EntireColumn);
            //hist.refer 1st copy
            rangeCc = xlWorkSheet2.UsedRange.get_Range(colC + "1", xlWorkSheet2.Cells[RowC, 3]) as Excel.Range;
            rangeCc2 = xlWorkSheet2.UsedRange.get_Range(colD + "1", xlWorkSheet2.Cells[RowC, 4]) as Excel.Range;
            rangeCc3 = xlWorkSheet2.UsedRange.get_Range(colE + "1", xlWorkSheet2.Cells[RowC, 5]) as Excel.Range;
            rangeCc3.Copy(rangeCc2);
            rangeCc3.Copy(rangeCc);
            //hist.preface 1st copy
            Excel.Range rangeOrigin1 = xlWorkSheetpre.Cells[1, 6] as Excel.Range;
            Excel.Range rangeMiddle1 = xlWorkSheetpre.Cells[1, 5] as Excel.Range;
            Excel.Range rangeReplace1 = xlWorkSheetpre.Cells[1, 4] as Excel.Range;
            rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
        //    rangeReplace1.EntireColumn.Replace("Historique!C", "Historique!E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
         //   rangeReplace1.EntireColumn.Replace("Hist.Refer!A", "Hist.Refer!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace1.EntireColumn.Copy(rangeOrigin1.EntireColumn);
            rangeReplace1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            //hist.calculs 1st copy
            rangexh1 = WorkSheetCalculs.Cells[1, 4] as Excel.Range;
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 3] as Excel.Range);
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 2] as Excel.Range);


            //hist 2ed copy
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            xlWorkSheet.get_Range("H1", "H1").EntireColumn.Copy(xlWorkSheet.get_Range("F1", "F1").EntireColumn);
            colcorrectif1 = cf.FindCodedColumnHeader("82000-1000", range);
            colcorrectif2 = cf.FindCodedColumnHeader("82000-2000", range);
            rangex1c = xlWorkSheet.UsedRange.get_Range(colcorrectif1 + "1", colcorrectif2 + "1") as Excel.Range;
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range(insert2 + "1", insert22 + "1").EntireColumn);
            //hist.refer 2ed copy
            rangeDc = xlWorkSheet2.UsedRange.get_Range(colF + "1", xlWorkSheet2.Cells[RowC, 6]) as Excel.Range;
            rangeDc2 = xlWorkSheet2.UsedRange.get_Range(colG + "1", xlWorkSheet2.Cells[RowC, 7]) as Excel.Range;
            rangeDc3 = xlWorkSheet2.UsedRange.get_Range(colH + "1", xlWorkSheet2.Cells[RowC, 8]) as Excel.Range;
            rangeDc3.Copy(rangeDc2);
            rangeDc3.Copy(rangeDc);
            //hist.preface 2ed copy
            Excel.Range rangeOrigin2 = xlWorkSheetpre.Cells[1, 10] as Excel.Range;
            Excel.Range rangeMiddle2 = xlWorkSheetpre.Cells[1, 9] as Excel.Range;
            Excel.Range rangeReplace2 = xlWorkSheetpre.Cells[1, 8] as Excel.Range;
            rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
       //     rangeReplace2.EntireColumn.Replace("Historique!F", "Historique!H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
      //      rangeReplace2.EntireColumn.Replace("Hist.Refer!D", "Hist.Refer!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace2.EntireColumn.Copy(rangeOrigin2.EntireColumn);
            rangeReplace2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            //hist.calculs 2ed copy
            rangexh1 = WorkSheetCalculs.Cells[1, 8] as Excel.Range;
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 7] as Excel.Range);
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 6] as Excel.Range);

            
          
           

            //hist 3ed copy
            xlWorkSheet.get_Range("K1", "K1").EntireColumn.Copy(xlWorkSheet.get_Range("I1", "I1").EntireColumn);
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            colcorrectif1 = cf.FindCodedColumnHeader("82000-1000", range);
            colcorrectif2 = cf.FindCodedColumnHeader("82000-2000", range);
            rangex1c = xlWorkSheet.UsedRange.get_Range(colcorrectif1 + "1", colcorrectif2 + "1") as Excel.Range;
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range(insert3+"1", insert32+"1").EntireColumn);
            //hist.refer 3ed copy

            rangeEc = xlWorkSheet2.UsedRange.get_Range(colI + "1", xlWorkSheet2.Cells[RowC, 9]) as Excel.Range;
            rangeEc2 = xlWorkSheet2.UsedRange.get_Range(colJ + "1", xlWorkSheet2.Cells[RowC, 10]) as Excel.Range;
            rangeEc3 = xlWorkSheet2.UsedRange.get_Range(colK + "1", xlWorkSheet2.Cells[RowC, 11]) as Excel.Range;
            rangeEc3.Copy(rangeEc2);
            rangeEc3.Copy(rangeEc);
            //hist.preface 3ed copy
            Excel.Range rangeOrigin3 = xlWorkSheetpre.Cells[1, 14] as Excel.Range;
            Excel.Range rangeMiddle3 = xlWorkSheetpre.Cells[1, 13] as Excel.Range;
            Excel.Range rangeReplace3 = xlWorkSheetpre.Cells[1, 12] as Excel.Range;
            rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
          //  rangeReplace3.EntireColumn.Replace("Historique!I", "Historique!K", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
         //   rangeReplace3.EntireColumn.Replace("Hist.Refer!G", "Hist.Refer!I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace3.EntireColumn.Copy(rangeOrigin3.EntireColumn);
            rangeReplace3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
            //hist.calculs 3ed copy
            rangexh1 = WorkSheetCalculs.Cells[1, 12] as Excel.Range;
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 11] as Excel.Range);
            rangexh1.EntireColumn.Copy(WorkSheetCalculs.Cells[1, 10] as Excel.Range);

            //alex:new two columns
            Excel.Range rangex4 = xlWorkSheet.Cells[1, 3] as Excel.Range;
            string insert41 = cf.FindCodedColumnHeader("3000", range);

            rangex4.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            Excel.Range rangex5 = xlWorkSheet.Cells[1, 3] as Excel.Range;
            string insert51 = cf.FindCodedColumnHeader("3000", range);
            rangex5.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            string insert4 = cf.FindCodedColumnHeader("83000-1000", range);
            string insert5 = cf.FindCodedColumnHeader("83000-2000", range);
            Excel.Range rangex2c = xlWorkSheet.UsedRange.get_Range(insert4 + "1", insert5 + "1") as Excel.Range;
            rangex2c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range(insert41 + "1", insert51 + "1").EntireColumn);

            //rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("D1", "E1").EntireColumn);
            //rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("G1", "H1").EntireColumn);
            //rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("J1", "K1").EntireColumn);


         
            //Excel.Range rangeC = xlWorkSheet2.Cells[1, 4] as Excel.Range;

            //Excel.Range rangeD = xlWorkSheet2.Cells[1, 5] as Excel.Range;

            //Excel.Range rangeE = xlWorkSheet2.Cells[1, 6] as Excel.Range;

            //rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
          


          

           

            
           



            Excel.Range rangex1cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 3] as Excel.Range;
            Excel.Range rangex1cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 4] as Excel.Range;
            Excel.Range rangex2cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 6] as Excel.Range;
            Excel.Range rangex2cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 7] as Excel.Range;
            Excel.Range rangex3cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 9] as Excel.Range;
            Excel.Range rangex3cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 10] as Excel.Range;
            Excel.Range rangex4cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 12] as Excel.Range;
            Excel.Range rangex4cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 13] as Excel.Range;
            rangex1cx1.Value2 = "";
            rangex1cx2.Value2 = "";
            rangex2cx1.Value2 = "";
            rangex2cx2.Value2 = "";
            rangex3cx1.Value2 = "";
            rangex3cx2.Value2 = "";
            rangex4cx1.Value2 = "";
            rangex4cx2.Value2 = "";




            //tester EE pour Histo.refer//et parcourir Historique

            Excel.Range rangeRefer = xlWorkSheet2.UsedRange;
            Excel.Range rangeHistorique = xlWorkSheet.UsedRange;
            //petite corr
            object[,] valuesRefer = (object[,])rangeRefer.Value2;
            object[,] valuesHistorique = (object[,])rangeHistorique.Value2;
            int rowCnt = 0;
            int rowHistoCnt = 0;
            string nomCol = "";

            //for (rowCnt = 1; rowCnt <= rangeRefer.Rows.Count; rowCnt++)
            //{
            //    string valuecellabs = Convert.ToString(valuesRefer[rowCnt, 1]);
            //    if (valuecellabs != "" && valuecellabs != "D" && valuecellabs != "D1" && valuecellabs != "d")
            //    {
            //        nomCol = valuecellabs;
            //        for (rowHistoCnt = 1; rowHistoCnt <= rangeHistorique.Rows.Count; rowHistoCnt++)
            //        {
            //            string valuecellHisto = Convert.ToString(valuesHistorique[rowHistoCnt, 2]);
            //            if (valuecellHisto == nomCol)
            //            {
            //                Excel.Range cellcopie = xlWorkSheet.Cells[rowHistoCnt, 5] as Excel.Range;
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 6]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 7]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 8]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 9]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 10]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 11]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 12]);
            //                cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 13]);
            //            }
            //        }
            //    }
            //}







        


            ///////////////////////////////Parcourir toutes les cellules ergodiaue////////////////////////////////
            //string str;
            //int rCnt = 0;
            //int cCnt = 0;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //    {
            //        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString();
            //        MessageBox.Show(str);
            //    }
            //}
            ///////////////////////////////fermer EXCEL automatiquement apres modification?//////////////////////
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //MessageBox.Show("jobs done!");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// supprimer colonne marqué "-1" Histo6.xls 944000
        //
        private void supprimercol_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////////////944000//////////////////////
            int row944000 = 0;

            //fonction pour trouver le numéro de ligne 944000
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row944000 = cf.FindCodedRow("944000", range);




            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row944000, col]);
                if (Regex.Equals(value, "-1"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row944000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// supprimer ligne marqué "-1" et "-2" forcage pour NP
        //
        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts =false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");

            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;
            int rCnt = 0;
            int cCnt = 0;
            int col = 0;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col = cf.FindCodedColumn("8000", range);
            rCnt = range.Rows.Count;


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "8000"))
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}
            int time1 = System.Environment.TickCount;

            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "-1"))//pour -1, -2 ensemble  || Regex.Equals(value, "-2")
                {
                    //MessageBox.Show(row.ToString());
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    rangeDelx.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row--;
                }
            }

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }
        //
        //// supprimer sauf langues pour nota-pme
        //
        private void supprimerhistosauflangues_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;

            Excel.Worksheet sheetpreface = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet sheetCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Worksheet sheetMacros = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Macros");
            Excel.Worksheet sheetCombos = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Combos et listes à cocher");
            sheetpreface.Delete();
            sheetCalculs.Delete();
            sheetMacros.Delete();
            sheetCombos.Delete();

            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// supprimer colonne marqué "-2" Histo8.xls avant diviser petite fichier pour Nota-pme
        //
        private void supprimermoin2_Click(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////944000//////////////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row944000 = 0;

            cCnt = range.Columns.Count;
            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "944000"))
                {
                    row944000 = rCnt;
                    break;
                }
            }

            //for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            //{
            //    string value = Convert.ToString(values[row944000, col]);
            //    if (Regex.Equals(value, "-2"))
            //    {
            //        Excel.Range rangeDelx = xlWorkSheet.Cells[row944000, col] as Excel.Range;
                    
            //        rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

            //        range = xlWorkSheet.UsedRange;
            //        values = (object[,])range.Value2;
            //        col--;
            //    }
            //}
            int nubmer=0;
            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row944000, col]);
                
                if (Regex.Equals(value, "-2"))
                {
                    nubmer ++;
                }
                else{
                    if (nubmer != 0)
                    {
                        Excel.Range rangeDelx = xlWorkSheet.get_Range(xlWorkSheet.Cells[row944000, col-nubmer],xlWorkSheet.Cells[row944000, col-1]) as Excel.Range;

                        rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                        range = xlWorkSheet.UsedRange;
                        values = (object[,])range.Value2;
                        col= col - nubmer;
                    }
                    nubmer = 0;
                }
            }
            range = xlWorkSheet.UsedRange;
            cCnt = range.Columns.Count;
            values = (object[,])range.Value2;
            for (int col = 1; col <= cCnt; col++)
            {
                string value = Convert.ToString(values[row944000, col]);
                if (Regex.Equals(value, "-4"))
                {
                    Excel.Range rangeEffacer = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col], xlWorkSheet.Cells[row944000-1, col]) as Excel.Range;
                    rangeEffacer.ClearContents();
                }
            }



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Histo.preface
        //
        private void Histopreface_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
           
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets["Hist.Preface"] as Excel.Worksheet;


           // CodeFinder cfcol;
           // cfcol = new CodeFinder(xlWorkBook, xlWorkSheet);

           // //l'ordre de déclaration à respecter
           // string colD = cfcol.FindCodedColumnHeader("4000", xlWorkSheet.UsedRange);
           // string colE = cfcol.FindCodedColumnHeader("5000", xlWorkSheet.UsedRange);
           // xlWorkBook.Save();
           // Excel.Range rangeinsert1 = xlWorkSheet.UsedRange.get_Range(colD+"1", colE+"1") as Excel.Range;
           // rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
           //// xlWorkBook.Save();
            

           // string colH = cfcol.FindCodedColumnHeader("6000", xlWorkSheet.UsedRange);
           // string colI = cfcol.FindCodedColumnHeader("7000", xlWorkSheet.UsedRange);
           // Excel.Range rangeinsert2 = xlWorkSheet.UsedRange.get_Range(colH+"1", colI+"1") as Excel.Range;
           // rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
           //// xlWorkBook.Save();


           // string colL = cfcol.FindCodedColumnHeader("8000", xlWorkSheet.UsedRange);
           // string colM = cfcol.FindCodedColumnHeader("9000", xlWorkSheet.UsedRange);
           // Excel.Range rangeinsert3 = xlWorkSheet.UsedRange.get_Range(colL+"1", colM+"1") as Excel.Range;
           // rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
           //// xlWorkBook.Save();
           // releaseObject(rangeinsert1);
           // releaseObject(rangeinsert2);
           // releaseObject(rangeinsert3);

           // //1
           // Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
           // Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
           // Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
           // rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
           // rangeReplace1.EntireColumn.Replace("Historique!C", "Historique!E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace1.EntireColumn.Replace("Hist.Refer!A", "Hist.Refer!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace1.EntireColumn.Copy(rangeOrigin1.EntireColumn);
           // rangeReplace1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
           // releaseObject(rangeOrigin1);
           // releaseObject(rangeMiddle1);
           // releaseObject(rangeReplace1);

           // //2
           // Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
           // Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
           // Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
           // rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
           // rangeReplace2.EntireColumn.Replace("Historique!F", "Historique!H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace2.EntireColumn.Replace("Hist.Refer!D", "Hist.Refer!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace2.EntireColumn.Copy(rangeOrigin2.EntireColumn);
           // rangeReplace2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
           // releaseObject(rangeOrigin2);
           // releaseObject(rangeMiddle2);
           // releaseObject(rangeReplace2);

           // //3
           // Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
           // Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
           // Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
           // rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
           // rangeReplace3.EntireColumn.Replace("Historique!I", "Historique!K", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace3.EntireColumn.Replace("Hist.Refer!G", "Hist.Refer!I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeReplace3.EntireColumn.Copy(rangeOrigin3.EntireColumn);
           // rangeReplace3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
           // releaseObject(rangeOrigin3);
           // releaseObject(rangeMiddle3);
           // releaseObject(rangeReplace3);

            Excel.Range rangeRef = xlWorkSheet.UsedRange;

            object[,] values = (object[,])rangeRef.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int Row500000 = 0;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            Row500000 = cf.FindCodedRow("500000", rangeRef);


            //cCnt = rangeRef.Columns.Count;
            //for (rCnt = 1; rCnt <= rangeRef.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "500000"))
            //    {
            //        Row500000 = rCnt;
            //        break;
            //    }
            //}
            Excel.Range rangeXLReplace1 = xlWorkSheet.Cells[Row500000, 8] as Excel.Range;
            Excel.Range rangeXLC11 = xlWorkSheet.Cells[Row500000, 9] as Excel.Range;
            Excel.Range rangeXLC12 = xlWorkSheet.Cells[Row500000, 10] as Excel.Range;
            //eviter bug vsto
            xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[Row500000, 8], xlWorkSheet.Cells[Row500000, 9]).Replace("Historique!A", "Historique!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeXLReplace1.Copy(rangeXLC11);
            rangeXLReplace1.Copy(rangeXLC12);
             

            Excel.Range rangeXLReplace2 = xlWorkSheet.Cells[Row500000, 12] as Excel.Range;
            Excel.Range rangeXLC21 = xlWorkSheet.Cells[Row500000, 13] as Excel.Range;
            Excel.Range rangeXLC22 = xlWorkSheet.Cells[Row500000, 14] as Excel.Range;
            //eviter bug vsto
            xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[Row500000, 12], xlWorkSheet.Cells[Row500000, 13]).Replace("Historique!D", "Historique!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeXLReplace2.Copy(rangeXLC21);
            rangeXLReplace2.Copy(rangeXLC22);



            //xlWorkBook.SaveCopyAs(prefaceNP);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Historique 84000 CF copie coller //supprimer ref! sur Hist.Refer!$c$8 //et ref! dans Hist.Refer
        //
        private void Historique84000()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet xlWorkSheetPreface = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet xlWorkSheetCalcul = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");

            Excel.Worksheet xlWorkSheetRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            Excel.Range range = xlWorkSheet.UsedRange;
            Excel.Range rangeRefer = xlWorkSheetRefer.UsedRange;
            Excel.Range rangePreface = xlWorkSheetPreface.UsedRange;
            Excel.Range rangeCalcul = xlWorkSheetCalcul.UsedRange;
            object[,] values = (object[,])range.Value2;
            int col = 0;

            //fonction pour trouver le numéro de col "84000"
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col = cf.FindCodedColumn("84000", range);


            int rCnt = 0;
            int cCnt = 0;
            rCnt = range.Rows.Count;
            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "84000"))
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}

            for (int i = 0; i < 2; i++)
            {
                //supprimer ref! dans Historique -- Hist.Refer!$E$3
                range.Cells.Replace("Historique!#REF!~*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!~*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!~*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                range.Cells.Replace("Historique!#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                range.Cells.Replace("Historique!#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("Historique!#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                range.Cells.Replace("#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                range.Cells.Replace("#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                range.Cells.Replace("#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);



                range.Cells.Replace("Hist.Refer!$Q$18", "0,005", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                //supprimer ref! dans Hist.Refer -- Hist.Refer!$E$3
                rangeRefer.Cells.Replace("Historique!#REF!~*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!~*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!~*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeRefer.Cells.Replace("Historique!#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeRefer.Cells.Replace("Historique!#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("Historique!#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeRefer.Cells.Replace("#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeRefer.Cells.Replace("#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeRefer.Cells.Replace("#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                //supprimer ref! dans Hist.preface
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$C$3", "0*Hist.Refer!$C$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$D$3", "0*Hist.Refer!$D$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$F$3", "0*Hist.Refer!$F$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$G$3", "0*Hist.Refer!$G$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$I$3", "0*Hist.Refer!$I$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$J$3", "0*Hist.Refer!$J$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!~*Hist.Refer!$K$8", "0*Hist.Refer!$K$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$C$3", "0*Hist.Refer!$C$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$D$3", "0*Hist.Refer!$D$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$F$3", "0*Hist.Refer!$F$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$G$3", "0*Hist.Refer!$G$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$I$3", "0*Hist.Refer!$I$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$J$3", "0*Hist.Refer!$J$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!*Hist.Refer!$K$8", "0*Hist.Refer!$K$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$C$3", "0)*Hist.Refer!$C$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$D$3", "0)*Hist.Refer!$D$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$F$3", "0)*Hist.Refer!$F$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$G$3", "0)*Hist.Refer!$G$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$I$3", "0)*Hist.Refer!$I$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$J$3", "0)*Hist.Refer!$J$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("Historique!#REF!)*Hist.Refer!$K$8", "0)*Hist.Refer!$K$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$C$3", "0)*Hist.Refer!$C$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$D$3", "0)*Hist.Refer!$D$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$F$3", "0)*Hist.Refer!$F$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$G$3", "0)*Hist.Refer!$G$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$I$3", "0)*Hist.Refer!$I$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$J$3", "0)*Hist.Refer!$J$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!)*Hist.Refer!$K$8", "0)*Hist.Refer!$K$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangePreface.Cells.Replace("#REF!*Hist.Refer!$C$3", "0*Hist.Refer!$C$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$D$3", "0*Hist.Refer!$D$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$F$3", "0*Hist.Refer!$F$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$G$3", "0*Hist.Refer!$G$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$I$3", "0*Hist.Refer!$I$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$J$3", "0*Hist.Refer!$J$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangePreface.Cells.Replace("#REF!*Hist.Refer!$K$8", "0*Hist.Refer!$K$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                //supprimer ref! dans Hist.Cacul
                rangeCalcul.Cells.Replace("Historique!#REF!~*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!~*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!~*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeCalcul.Cells.Replace("Historique!#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeCalcul.Cells.Replace("Historique!#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("Historique!#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeCalcul.Cells.Replace("#REF!*Hist.Refer!$E$3", "0*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("#REF!*Hist.Refer!$H$3", "0*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("#REF!*Hist.Refer!$K$3", "0*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                rangeCalcul.Cells.Replace("#REF!)*Hist.Refer!$E$3", "0)*Hist.Refer!$E$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("#REF!)*Hist.Refer!$H$3", "0)*Hist.Refer!$H$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                rangeCalcul.Cells.Replace("#REF!)*Hist.Refer!$K$3", "0)*Hist.Refer!$K$3", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                Excel.Range rangecopie = xlWorkSheet.Cells[2, col] as Excel.Range;
                Excel.Range rangecoller = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[2, col], xlWorkSheet.Cells[rCnt - 4, col]) as Excel.Range;
            }
           // rangecoller.Copy(misValue);
            //rangecopie.Copy(rangecoller);
            
           
            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkSheetPreface);
            releaseObject(xlWorkSheetCalcul);
            releaseObject(xlWorkSheetRefer);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Historique Col xxxxx consigne de cellule proteger
        //
        private void consigneProteger()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            int rowcount = xlWorkSheet.UsedRange.Rows.Count;
            object[,] values = (object[,])range.Value2;
            int col = 0;

            //fonction pour trouver le numéro de col
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col = cf.FindCodedColumn("15000", range);

            //int rCnt = 0;
            //int cCnt = 0;
            
            //rCnt = range.Rows.Count;
            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "15000"))
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}

            //Routine pour modifier col XXXXX marquer ligne proteger -1
            for (int i = 1; i < rowcount-5; i++)
            {
                if ((xlWorkSheet.Cells[i, 5] as Excel.Range).Locked.ToString() == "True")
                    (xlWorkSheet.Cells[i, col] as Excel.Range).Value2 = "-1";
            }


            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Hist.Refer pour D1 les formule trop long
        //
        private void fonctionRemplacerD1()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            xlWorkSheet.Name = "B";
            Excel.Worksheet xlWorkSheetB = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("B");
            Excel.Range rangeRefer = xlWorkSheetB.UsedRange;

            object[,] formuleRefer = (object[,])rangeRefer.Formula;


            int rowCnt = 0;

            for (rowCnt = 1; rowCnt <= rangeRefer.Rows.Count; rowCnt++)
            {
                string valuecellabs = Convert.ToString((xlWorkSheetB.Cells[rowCnt, 1] as Excel.Range).Value2);
                if (valuecellabs == "D1")
                {
                    Excel.Range rangeRep = xlWorkSheet.Cells[rowCnt, 1] as Excel.Range;
                    rangeRep.EntireRow.Cells.Replace("Historique!#REF!~*B!$C$8", "0*B!$C$8", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                }
            }

            xlWorkSheetB.Name = "Hist.Refer";
            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Changer tous les premiere numero de ligne, 1000 par 1000-100 xlWorkSheet ParamSav 0 1
        //
        private void changerNumeroligne()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues");
            Excel.Worksheet xlWorkSheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            int col = xlWorkSheet.UsedRange.Columns.Count;
            int col2 = xlWorkSheet2.UsedRange.Columns.Count;
            int col3 = xlWorkSheet3.UsedRange.Columns.Count;
            Excel.Range rangeChanger = xlWorkSheet.UsedRange.Cells[1, col] as Excel.Range;
            Excel.Range rangeChanger2 = xlWorkSheet2.UsedRange.Cells[1, col2] as Excel.Range;
            Excel.Range rangeChanger3 = xlWorkSheet3.UsedRange.Cells[1, col3] as Excel.Range;
            rangeChanger.Copy();
            rangeChanger2.Copy();
            rangeChanger3.Copy();
            rangeChanger.Value2 = "1000-100";
            rangeChanger2.Value2 = "1000-100";
            rangeChanger3.Value2 = "1000-100";


            //traitement pour choix, le value de sheet, // param sav
            Excel.Worksheet xlWorkSheetParamSav = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");

            Excel.Range range = xlWorkSheetParamSav.UsedRange;
            object[,] values = (object[,])range.Value2;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheetParamSav);


            int rCnt = 0;
            int cCnt = 0;
            int colx = 0;
            int row267000 = 0;
            int row268000 = 0;

            colx = cf.FindCodedColumn("3000", range);
            rCnt = range.Rows.Count;
            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "3000"))
            //    {
            //        colx = cCnt;
            //        break;
            //    }
            //}

            row267000 = cf.FindCodedRow("267000", range);
            row268000 = cf.FindCodedRow("268000", range);
            cCnt = range.Columns.Count;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "267000"))
            //    {
            //        row267000 = rCnt;
            //    }

            //    if (Regex.Equals(valuecellabs, "268000"))
            //    {
            //        row268000 = rCnt;
            //        break;
            //    }
            //}



            Excel.Range rangecell267000 = xlWorkSheetParamSav.Cells[row267000, colx] as Excel.Range;
            Excel.Range rangecell268000 = xlWorkSheetParamSav.Cells[row268000, colx] as Excel.Range;
            rangecell267000.Value2 = 0;
            rangecell268000.Value2 = 1;



            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkSheet2);
            releaseObject(xlWorkSheet3);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Annuel.ptw  "O"    col 33000 copie coller
        //
        private void AnnuelO_Click(object sender, EventArgs e)
        {
           
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook0;
            Excel.Workbook xlWorkBook;//Annuel.ptw
            Excel.Workbook xlWorkBook2;//Admin.ptw
            Excel.Workbook xlWorkBook3;//Histo.ptw
            Excel.Workbook xlWorkBook4;//Eval.ptw
            Excel.Workbook xlWorkBook5;//Decis.ptw
            Excel.Workbook xlWorkBook6;//Tres.ptw
            Excel.Workbook xlWorkBook7;//Histo-s.ptw
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            
            //fichier intermedia annuel.xls //peut effacer apres

            //xlWorkBook0 = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook0.SaveCopyAs("D:\\ptw\\Annuel.xls");
            //xlWorkBook0.Close(true, misValue, misValue);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook0 = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkBook0.SaveCopyAs("D:\\ptw\\Annuel.xlsx");
            xlWorkBook0.Close(true, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //xlWorkBook = xlApp.Workbooks.Open("D:\\ptw\\Annuel.xls", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook2 = xlApp.Workbooks.Open("D:\\ptw\\Admin.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook3 = xlApp.Workbooks.Open("D:\\ptw\\Histo.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook4 = xlApp.Workbooks.Open("D:\\ptw\\Eval.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook5 = xlApp.Workbooks.Open("D:\\ptw\\Decis.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook6 = xlApp.Workbooks.Open("D:\\ptw\\Tres.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook7 = xlApp.Workbooks.Open("D:\\ptw\\Histo-s.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Range range = xlWorkSheet.UsedRange;
            Excel.Worksheet xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range rangeCompAnn = xlWorkSheet2.UsedRange;
            Excel._Worksheet xlWorksheet3 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            // Excel.Range rangeAnnu = xlWorksheet3.UsedRange;
            Excel.Worksheet xlWorkSheet4 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
         
            //////////////////////////////supprimer -6 AL "19000-1000"
            object[,] values = (object[,])rangeCompAnn.Value2;
            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int rnumb = 0;
            rCnt = rangeCompAnn.Rows.Count;
            for (cCnt = 1; cCnt <= rangeCompAnn.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "19000-1000"))
                {
                    col = cCnt;
                    break;
                }
                
            }

            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "-6"))
                {
                    //MessageBox.Show(row.ToString());
                    Excel.Range rangeDelx = xlWorkSheet2.Cells[row, col] as Excel.Range;
                    rangeDelx.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    rangeCompAnn = xlWorkSheet2.UsedRange;
                    values = (object[,])rangeCompAnn.Value2;
                    row--;
                }
                string valuex = Convert.ToString(values[row, rangeCompAnn.Columns.Count]);
                if (Regex.Equals(valuex, "539000-3000"))
                {
                    rnumb = row;
                   
                }
            }
            ///////////////////////////////////Onglet O//////////////////////////////////////
            Excel.Range rangeinsert1 = xlWorkSheet.UsedRange.get_Range("D1", "E1") as Excel.Range;
            rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangeinsert2 = xlWorkSheet.UsedRange.get_Range("H1", "I1") as Excel.Range;
            rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangeinsert3 = xlWorkSheet.UsedRange.get_Range("L1", "M1") as Excel.Range;
            rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            ///////////////////////////////Onglet Comptes Annuels////////////////////////////

            //alex change
            Excel.Range rangeinsert12 = xlWorkSheet2.UsedRange.get_Range("D1", "E1") as Excel.Range;
            rangeinsert12.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
          //  xlWorkBook.Save();
            Excel.Range rangeinsert22 = xlWorkSheet2.UsedRange.get_Range("H1", "I1") as Excel.Range;
            rangeinsert22.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
           // xlWorkBook.Save();
            Excel.Range rangeinsert32 = xlWorkSheet2.UsedRange.get_Range("L1", "M1") as Excel.Range;
            rangeinsert32.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
          //  xlWorkBook.Save();

         

            Excel.Range rangerefer1 = xlWorksheet3.UsedRange.get_Range("D1", "E1") as Excel.Range;
            rangerefer1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            Excel.Range rangerefer2 = xlWorksheet3.UsedRange.get_Range("H1", "I1") as Excel.Range;
            rangerefer2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            Excel.Range rangerefer3 = xlWorksheet3.UsedRange.get_Range("L1", "M1") as Excel.Range;
            rangerefer3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            
           
            //1
            Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
            Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
            Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
            rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
            rangeOrigin1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            //2
            Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
            Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
            Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
            rangeOrigin2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            //3
            Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
            Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
            Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
            rangeOrigin3.EntireColumn.Copy(rangeMiddle3.EntireColumn);

            rangeinsert1 = xlWorkSheet.UsedRange.get_Range("H1", "J1") as Excel.Range;
            rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeinsert1 = xlWorkSheet.UsedRange.get_Range("H1") as Excel.Range;
            //rangeinsert1.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            rangeinsert2 = xlWorkSheet.UsedRange.get_Range("O1", "Q1") as Excel.Range;
            rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeinsert2 = xlWorkSheet.UsedRange.get_Range("O1") as Excel.Range;
            //rangeinsert2.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangeinsert3 = xlWorkSheet.UsedRange.get_Range("V1", " X1") as Excel.Range;
            rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
                       //rangeinsert3 = xlWorkSheet.UsedRange.get_Range("V1") as Excel.Range;
           // rangeinsert3.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            ///////////////////////////////////Onglet O//////////////////////////////////////
            ///////////////////////////////Onglet Comptes Annuels////////////////////////////
            //1
          





            Excel.Range rangeOrigin12 = xlWorkSheet2.Cells[1, 6] as Excel.Range;
            rangeOrigin12.EntireColumn.Replace("Historique!E", "Historique!G", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           
            Excel.Range r1 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[rnumb, 5],xlWorkSheet2.Cells[rnumb, 6]) as Excel.Range;
          //  object[,] xxxx= (object[,])r1.Formula;
          //  string xxx = xxxx[1,1].ToString();
           
            r1.Cells.Replace("C", "E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByColumns, false, Type.Missing, false, false);
            //rangeOrigin12.EntireColumn.Replace("C737*C658", "E737*E658", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            Excel.Range rangeMiddle12 = xlWorkSheet2.Cells[1, 5] as Excel.Range;
            Excel.Range rangeReplace12 = xlWorkSheet2.Cells[1, 4] as Excel.Range;
            rangeOrigin12.EntireColumn.Copy(rangeMiddle12.EntireColumn);
            rangeOrigin12.EntireColumn.Copy(rangeReplace12.EntireColumn);
            //2
            Excel.Range rangeOrigin22 = xlWorkSheet2.Cells[1, 10] as Excel.Range;
            rangeOrigin22.EntireColumn.Replace("Historique!H", "Historique!J", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
          
            Excel.Range r2 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[rnumb, 9], xlWorkSheet2.Cells[rnumb, 10]) as Excel.Range;
            r2.Replace("G", "I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeOrigin12.EntireColumn.Replace("G737*G658", "L737*L658", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            Excel.Range rangeMiddle22 = xlWorkSheet2.Cells[1, 9] as Excel.Range;
            Excel.Range rangeReplace22 = xlWorkSheet2.Cells[1, 8] as Excel.Range;
            rangeOrigin22.EntireColumn.Copy(rangeMiddle22.EntireColumn);
            rangeOrigin22.EntireColumn.Copy(rangeReplace22.EntireColumn);
            //3
            Excel.Range rangeOrigin32 = xlWorkSheet2.Cells[1, 14] as Excel.Range;
            rangeOrigin32.EntireColumn.Replace("Historique!K", "Historique!M", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

           
            Excel.Range r3 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[rnumb, 13], xlWorkSheet2.Cells[rnumb, 14]) as Excel.Range;
            r3.Replace("K", "M", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // rangeOrigin12.EntireColumn.Replace("N737*N658", "T737*T658", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            Excel.Range rangeMiddle32 = xlWorkSheet2.Cells[1, 13] as Excel.Range;
            Excel.Range rangeReplace32 = xlWorkSheet2.Cells[1, 12] as Excel.Range;
            rangeOrigin32.EntireColumn.Copy(rangeMiddle32.EntireColumn);
            rangeOrigin32.EntireColumn.Copy(rangeReplace32.EntireColumn);
            //alex:test1
            Excel.Range rangeinsert4 = xlWorkSheet2.UsedRange.get_Range("H1", "J1") as Excel.Range;
            rangeinsert4.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangeinsert5 = xlWorkSheet2.UsedRange.get_Range("O1", "Q1") as Excel.Range;
            rangeinsert5.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangeinsert6 = xlWorkSheet2.UsedRange.get_Range("V1", "X1") as Excel.Range;
            rangeinsert6.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangerefer4 = xlWorksheet3.UsedRange.get_Range("H1", "J1") as Excel.Range;
            rangerefer4.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangerefer5 = xlWorksheet3.UsedRange.get_Range("O1", "Q1") as Excel.Range;
            rangerefer5.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangerefer6 = xlWorksheet3.UsedRange.get_Range("V1", "X1") as Excel.Range;
            rangerefer6.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range r4 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[rnumb, 10], xlWorkSheet2.Cells[rnumb, 11]) as Excel.Range;
            r4.Replace("G", "J", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            Excel.Range r5 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[rnumb, 17], xlWorkSheet2.Cells[rnumb, 18]) as Excel.Range;
            r5.Replace("N", "Q", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            /////////////////////////////// deuxieme insertion//////////////////////////////
            //alex:test1
            rangeinsert4 = xlWorkSheet2.UsedRange.get_Range("H1", "J1") as Excel.Range;
            rangeinsert5 = xlWorkSheet2.UsedRange.get_Range("O1", "Q1") as Excel.Range;
            rangeinsert6 = xlWorkSheet2.UsedRange.get_Range("V1", "X1") as Excel.Range;
            ////Alex:test2 and old version
            //Excel.Range range2ndInsertionX1 = xlWorkSheet2.Cells[1, 8] as Excel.Range;
            //Excel.Range range2ndInsertionX2 = xlWorkSheet2.Cells[1, 12] as Excel.Range;
            //Excel.Range range2ndInsertionX3 = xlWorkSheet2.Cells[1, 16] as Excel.Range;
            //alex:test1
            Excel.Range range2ndInsertion1 = xlWorkSheet2.UsedRange.get_Range("BW1", "BY1") as Excel.Range;
            //  range2ndInsertion1.EntireColumn.Replace("Annu.Refer!BH", "Annu.Refer!BW", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range2ndInsertion1.EntireColumn.Copy(rangeinsert4.EntireColumn);
            range2ndInsertion1.EntireColumn.Copy(rangeinsert5.EntireColumn);
            range2ndInsertion1.EntireColumn.Copy(rangeinsert6.EntireColumn);
            //alex: test2
            //Excel.Range range2ndInsertionreplace = xlWorkSheet2.UsedRange.get_Range("BN1", "BP1") as Excel.Range;
            //range2ndInsertionreplace.EntireColumn.Replace("Annu.Refer!BH", "Annu.Refer!BQ", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            //alex:test2 and old version
            //Excel.Range range2ndInsertion1 = xlWorkSheet2.UsedRange.get_Range("BN1", "BP1") as Excel.Range;
            //range2ndInsertion1.EntireColumn.Copy(misValue);
            //range2ndInsertionX1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            //Excel.Range range2ndInsertion2 = xlWorkSheet2.UsedRange.get_Range("BQ1", "BS1") as Excel.Range;
            //range2ndInsertion2.EntireColumn.Copy(misValue);
            //range2ndInsertionX2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            //Excel.Range range2ndInsertion3 = xlWorkSheet2.UsedRange.get_Range("BT1", "BV1") as Excel.Range;
            //range2ndInsertion3.EntireColumn.Copy(misValue);
            //range2ndInsertionX3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            int col33000 = 0;
            int col3000250 = 0;
            int col2000250 = 0;
            int col4000250 = 0;
            Excel.Range newrangeCompAnn = xlWorkSheet2.UsedRange;
            object[,] newvalues = (object[,])newrangeCompAnn.Value2;
            rCnt = newrangeCompAnn.Rows.Count;
            for (cCnt = 1; cCnt <= newrangeCompAnn.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(newvalues[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "33000"))
                {
                    col33000 = cCnt;
                    break;
                }
                if (Regex.Equals(valuecellabs, "3000-250"))
                {
                    col3000250 = cCnt;
                  
                }
                if (Regex.Equals(valuecellabs, "4000-250"))
                {
                    col4000250 = cCnt;
                    
                }
                if (Regex.Equals(valuecellabs, "2000-250"))
                {
                    col2000250 = cCnt;

                }
            }

            int row33000Cnt = rangeCompAnn.Rows.Count;
            //alex  sdfsdfsdfsdfsdfsdf

            //Alex change and replace the formula after all finished

            Excel.Range range3000250 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col3000250], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col3000250]) as Excel.Range;
            range3000250.EntireColumn.Replace("+($F", "+(I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range3000250.EntireColumn.Replace("+$G", "+J", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range3000250.EntireColumn.Replace("+$B", "+C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            range3000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col3000250+1], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col3000250+1]).EntireColumn);
            range3000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col3000250 + 2], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col3000250 + 2]).EntireColumn);


            Excel.Range range4000250 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col4000250], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col4000250]) as Excel.Range;
            range4000250.EntireColumn.Replace("+($M", "+(P", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range4000250.EntireColumn.Replace("+$N", "+Q", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range4000250.EntireColumn.Replace("+$C", "+J", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range4000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col4000250 + 1], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col4000250 + 1]).EntireColumn);
            range4000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col4000250 + 2], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col4000250 + 2]).EntireColumn);







            Excel.Range rangecopie = xlWorkSheet2.Cells[2, col33000] as Excel.Range;
            Excel.Range rangecoller = xlWorkSheet2.UsedRange.get_Range(xlWorkSheet2.Cells[2, col33000], xlWorkSheet2.Cells[row33000Cnt - 4, col33000]) as Excel.Range;
            rangecoller.Copy(misValue);
            rangecopie.Copy(rangecoller);

            Excel.Range rangex1cx1 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 7] as Excel.Range;
            Excel.Range rangex1cx2 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 8] as Excel.Range;
            Excel.Range rangex1cx3 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 9] as Excel.Range;
            Excel.Range rangex1cx4 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 10] as Excel.Range;
            Excel.Range rangex2cx1 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 14] as Excel.Range;
            Excel.Range rangex2cx2 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 15] as Excel.Range;
            Excel.Range rangex2cx3 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 16] as Excel.Range;
            Excel.Range rangex2cx4 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 17] as Excel.Range;
            Excel.Range rangex3cx1 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 21] as Excel.Range;
            Excel.Range rangex3cx2 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 22] as Excel.Range;
            Excel.Range rangex3cx3 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 23] as Excel.Range;
            Excel.Range rangex3cx4 = xlWorkSheet2.Cells[rangeCompAnn.Rows.Count - 1, 24] as Excel.Range;
            rangex1cx1.Value2 = "";
            rangex1cx2.Value2 = "";
            rangex2cx1.Value2 = "";
            rangex2cx2.Value2 = "";
            rangex3cx1.Value2 = "";
            rangex3cx2.Value2 = "";
            rangex1cx3.Value2 = "";
            rangex2cx3.Value2 = "";
            rangex3cx3.Value2 = "";
            rangex1cx4.Value2 = "";
            rangex2cx4.Value2 = "";
            rangex3cx4.Value2 = "";
            Excel.Range rngecopyx = xlWorkSheet.UsedRange.get_Range("AT1", "AV1") as Excel.Range;
            rngecopyx.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("H1", "J1"));
            rngecopyx.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("O1", "Q1"));
            rngecopyx.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("V1", "X1"));
           // Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            //Excel.Range range = xlWorkSheet.UsedRange;
            xlWorkSheet.UsedRange.get_Range("H1", "H848").EntireColumn.Replace("!$N", "!$H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("O1", "O848").EntireColumn.Replace("!$N", "!$O", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("V1", "V848").EntireColumn.Replace("!$N", "!$V", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            Excel.Range range2000250 = xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col2000250], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col2000250]) as Excel.Range;
            range2000250.EntireColumn.Replace("+($B", "+(B", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range2000250.EntireColumn.Replace("+$C", "+C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
           // range2000250.EntireColumn.Replace("+$C", "+J", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            range2000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col2000250 + 1], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col2000250 + 1]).EntireColumn);
            range2000250.EntireColumn.Copy(xlWorkSheet2.get_Range(xlWorkSheet2.Cells[2, col2000250 + 2], xlWorkSheet2.Cells[xlWorkSheet2.UsedRange.Rows.Count - 1, col2000250 + 2]).EntireColumn);

               Excel.Range rangeP = xlWorkSheet4.UsedRange;
               int rowcount = 291;
               for (int xcount = 0; xcount < 5; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!F$", "!D$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 296;
               for (int xcount = 0; xcount < 5; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!M$", "!K$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 301;
               for (int xcount = 0; xcount < 5; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!T$", "!R$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 710;
               for (int xcount = 0; xcount < 5; xcount++)
               {
                   
                   Excel.Range range710 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range710.EntireRow.Replace("!$U", "!$I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 715;
               for (int xcount = 0; xcount < 5; xcount++)
               {

                   Excel.Range range710 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range710.EntireRow.Replace("!$U", "!$P", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 720;
               for (int xcount = 0; xcount < 5; xcount++)
               {

                   Excel.Range range710 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range710.EntireRow.Replace("!$U", "!$W", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }

                rowcount = 383;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!F$", "!D$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 391;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!M$", "!K$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }

               rowcount = 399;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!T$", "!R$", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 734;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!$U", "!$I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 742;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!$U", "!$P", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
               rowcount = 750;
               for (int xcount = 0; xcount < 8; xcount++)
               {
                   Excel.Range range291 = xlWorkSheet4.get_Range(xlWorkSheet4.Cells[rowcount + xcount, 1], xlWorkSheet4.Cells[rowcount + xcount, rangeP.Columns.Count]) as Excel.Range;
                   range291.EntireRow.Replace("!$U", "!$W", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
               }
            xlWorkBook.SaveAs(prefaceNP);
            //xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkSheet2);
            releaseObject(xlWorkBook0);
            releaseObject(xlWorkBook);
            //releaseObject(xlWorkBook2);
            //releaseObject(xlWorkBook3);
            //releaseObject(xlWorkBook4);
            //releaseObject(xlWorkBook5);
            //releaseObject(xlWorkBook6);
            //releaseObject(xlWorkBook7);
            releaseObject(xlApp);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            //Excel.Range range = xlWorkSheet.UsedRange;
            //Excel.Worksheet xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            //Excel.Range rangeCompAnn = xlWorkSheet2.UsedRange;

            ////////////////////////////////supprimer -6 AL "19000-1000"
            //object[,] values = (object[,])rangeCompAnn.Value2;
            //int col = 0;
            //int rCnt = 0;
            //int cCnt = 0;
            ////fonction pour trouver le numéro de col "8000"
            //CodeFinder cf;
            //cf = new CodeFinder(xlWorkBook, xlWorkSheet2);
            //col = cf.FindCodedColumn("8000", rangeCompAnn);


            ////l'ancien fonction a supprimer



            ////rCnt = rangeCompAnn.Rows.Count;
            ////for (cCnt = 1; cCnt <= rangeCompAnn.Columns.Count; cCnt++)
            ////{
            ////    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            ////    if (Regex.Equals(valuecellabs, "19000-1000"))
            ////    {
            ////        col = cCnt;
            ////        break;
            ////    }
            ////}




            //for (int row = 1; row <= values.GetUpperBound(0); row++)
            //{
            //    string value = Convert.ToString(values[row, col]);
            //    if (Regex.Equals(value, "-6"))
            //    {
            //        //MessageBox.Show(row.ToString());
            //        Excel.Range rangeDelx = xlWorkSheet2.Cells[row, col] as Excel.Range;
            //        rangeDelx.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            //        rangeCompAnn = xlWorkSheet2.UsedRange;
            //        values = (object[,])rangeCompAnn.Value2;
            //        row--;
            //    }
            //}
            /////////////////////////////////Onglet Comptes Annuels////////////////////////////
            //Excel.Range rangeinsert12 = xlWorkSheet2.UsedRange.get_Range("D1", "E1") as Excel.Range;
            //rangeinsert12.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //xlWorkBook.Save();
            //Excel.Range rangeinsert22 = xlWorkSheet2.UsedRange.get_Range("H1", "I1") as Excel.Range;
            //rangeinsert22.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //xlWorkBook.Save();
            //Excel.Range rangeinsert32 = xlWorkSheet2.UsedRange.get_Range("L1", "M1") as Excel.Range;
            //rangeinsert32.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //xlWorkBook.Save();
            /////////////////////////////////////Onglet O//////////////////////////////////////
            //Excel.Range rangeinsert1 = xlWorkSheet.UsedRange.get_Range("D1", "E1") as Excel.Range;
            //rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //Excel.Range rangeinsert2 = xlWorkSheet.UsedRange.get_Range("H1", "I1") as Excel.Range;
            //rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //Excel.Range rangeinsert3 = xlWorkSheet.UsedRange.get_Range("L1", "M1") as Excel.Range;
            //rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            ////1
            //Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
            //Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
            //Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
            //rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
            //rangeOrigin1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            ////2
            //Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
            //Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
            //Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            //rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
            //rangeOrigin2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            ////3
            //Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
            //Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
            //Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            //rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
            //rangeOrigin3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
            /////////////////////////////////////Onglet O//////////////////////////////////////
            /////////////////////////////////Onglet Comptes Annuels////////////////////////////
            ////1
            //Excel.Range rangeOrigin12 = xlWorkSheet2.Cells[1, 6] as Excel.Range;
            //Excel.Range rangeMiddle12 = xlWorkSheet2.Cells[1, 5] as Excel.Range;
            //Excel.Range rangeReplace12 = xlWorkSheet2.Cells[1, 4] as Excel.Range;
            //rangeOrigin12.EntireColumn.Copy(rangeMiddle12.EntireColumn);
            //rangeOrigin12.EntireColumn.Copy(rangeReplace12.EntireColumn);
            ////2
            //Excel.Range rangeOrigin22 = xlWorkSheet2.Cells[1, 10] as Excel.Range;
            //Excel.Range rangeMiddle22 = xlWorkSheet2.Cells[1, 9] as Excel.Range;
            //Excel.Range rangeReplace22 = xlWorkSheet2.Cells[1, 8] as Excel.Range;
            //rangeOrigin22.EntireColumn.Copy(rangeMiddle22.EntireColumn);
            //rangeOrigin22.EntireColumn.Copy(rangeReplace22.EntireColumn);
            ////3
            //Excel.Range rangeOrigin32 = xlWorkSheet2.Cells[1, 14] as Excel.Range;
            //Excel.Range rangeMiddle32 = xlWorkSheet2.Cells[1, 13] as Excel.Range;
            //Excel.Range rangeReplace32 = xlWorkSheet2.Cells[1, 12] as Excel.Range;
            //rangeOrigin32.EntireColumn.Copy(rangeMiddle32.EntireColumn);
            //rangeOrigin32.EntireColumn.Copy(rangeReplace32.EntireColumn);

            ///////////////////////////////// deuxieme insertion//////////////////////////////
            //Excel.Range range2ndInsertionX1 = xlWorkSheet2.Cells[1, 8] as Excel.Range;
            //Excel.Range range2ndInsertionX2 = xlWorkSheet2.Cells[1, 12] as Excel.Range;
            //Excel.Range range2ndInsertionX3 = xlWorkSheet2.Cells[1, 16] as Excel.Range;

            //Excel.Range range2ndInsertion1 = xlWorkSheet2.UsedRange.get_Range("BN1", "BP1") as Excel.Range;
            //range2ndInsertion1.EntireColumn.Copy(misValue);
            //range2ndInsertionX1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            //Excel.Range range2ndInsertion2 = xlWorkSheet2.UsedRange.get_Range("BQ1", "BS1") as Excel.Range;
            //range2ndInsertion2.EntireColumn.Copy(misValue);
            //range2ndInsertionX2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            //Excel.Range range2ndInsertion3 = xlWorkSheet2.UsedRange.get_Range("BT1", "BV1") as Excel.Range;
            //range2ndInsertion3.EntireColumn.Copy(misValue);
            //range2ndInsertionX3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            //int col33000 = 0;
            //Excel.Range newrangeCompAnn = xlWorkSheet2.UsedRange;
            //object[,] newvalues = (object[,])newrangeCompAnn.Value2;


            //col33000 = cf.FindCodedColumn("33000", rangeCompAnn);


            ////rCnt = newrangeCompAnn.Rows.Count;
            ////for (cCnt = 1; cCnt <= newrangeCompAnn.Columns.Count; cCnt++)
            ////{
            ////    string valuecellabs = Convert.ToString(newvalues[rCnt, cCnt]);
            ////    if (Regex.Equals(valuecellabs, "33000"))
            ////    {
            ////        col33000 = cCnt;
            ////        break;
            ////    }
            ////}

            //int row33000Cnt = rangeCompAnn.Rows.Count;

            //Excel.Range rangecopie = xlWorkSheet2.Cells[2, col33000] as Excel.Range;
            //Excel.Range rangecoller = xlWorkSheet2.UsedRange.get_Range(xlWorkSheet2.Cells[2, col33000], xlWorkSheet2.Cells[row33000Cnt - 4, col33000]) as Excel.Range;
            //rangecoller.Copy(misValue);
            //rangecopie.Copy(rangecoller);


            //xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();


            //releaseObject(xlWorkSheet);
            //releaseObject(xlWorkSheet2);
            //releaseObject(xlWorkBook);

            //releaseObject(xlApp);
           

           

        }
        //
        //// Alleger Annuel.ptw
        //
        private void AllegerAnnuel_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook0;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //xlWorkBook0 = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook0 = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);



            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item("Comptes annuels");

            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;
            int rCnt = 0;
            int cCnt = 0;
            int col11000 = 0;
            int row1012000 = 0;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook0, xlWorkSheet);


            col11000 = cf.FindCodedColumn("11000-1000", range);
            //rCnt = range.Rows.Count;
            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "11000-1000"))
            //    {
            //        col11000 = cCnt;
            //        break;
            //    }
            //}

            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col11000]);
                if (Regex.Equals(value, "-6"))
                {
                    //MessageBox.Show(row.ToString());
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col11000] as Excel.Range;
                    rangeDelx.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row--;
                }
            }


            row1012000 = cf.FindCodedRow("1012000", range);
            //cCnt = range.Columns.Count;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "1012000"))
            //    {
            //        row1012000 = rCnt;
            //        break;
            //    }
            //}

            ////////////////////////////////////////-2//////////////////////////////

            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row1012000, col]);
                if (Regex.Equals(value, "-2"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row1012000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook0.Close(true, misValue, misValue);
            xlApp.Save(misValue);
            xlApp.Quit();



            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook0);
            releaseObject(xlApp);
        }
        //
        //// Annuel.ptw  "Comptes Annuels"
        //
        private void ComptesAnnuels_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook0;


            Excel.Workbook xlWorkBook;//Annuel.ptw
            Excel.Workbook xlWorkBook2;//Admin.ptw
            Excel.Workbook xlWorkBook3;//Histo.ptw
            Excel.Workbook xlWorkBook4;//Eval.ptw
            Excel.Workbook xlWorkBook5;//Decis.ptw
            Excel.Workbook xlWorkBook6;//Tres.ptw
            Excel.Workbook xlWorkBook7;//Histo-s.ptw
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //xlWorkBook0 = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook0.SaveCopyAs("D:\\ptw\\Annuel.ptw");
            //xlWorkBook0.Close(true, misValue, misValue);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            

            //xlWorkBook = xlApp.Workbooks.Open("D:\\ptw\\Annuel.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook2 = xlApp.Workbooks.Open("D:\\ptw\\Admin.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook3 = xlApp.Workbooks.Open("D:\\ptw\\Histo.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook4 = xlApp.Workbooks.Open("D:\\ptw\\Eval.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook5 = xlApp.Workbooks.Open("D:\\ptw\\Decis.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook6 = xlApp.Workbooks.Open("D:\\ptw\\Tres.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //xlWorkBook7 = xlApp.Workbooks.Open("D:\\ptw\\Histo-s.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;


            ////////////////////////////////////////////// premiere insertion
            Excel.Range rangeinsert1 = xlWorkSheet.UsedRange.get_Range("D1", "E1") as Excel.Range;
            rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();
            Excel.Range rangeinsert2 = xlWorkSheet.UsedRange.get_Range("H1", "I1") as Excel.Range;
            rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();
            Excel.Range rangeinsert3 = xlWorkSheet.UsedRange.get_Range("L1", "M1") as Excel.Range;
            rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();

            //1
            Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
            Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
            Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
            rangeOrigin1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
            //2
            Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
            Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
            Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            rangeOrigin2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
            //3
            Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
            Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
            Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            rangeOrigin3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
            rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);

            ////////////////////////////////////////////// deuxieme insertion
            Excel.Range range2ndInsertionX1 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            Excel.Range range2ndInsertionX2 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            Excel.Range range2ndInsertionX3 = xlWorkSheet.Cells[1, 16] as Excel.Range;

            Excel.Range range2ndInsertion1 = xlWorkSheet.UsedRange.get_Range("BN1", "BP1") as Excel.Range;
            range2ndInsertion1.EntireColumn.Copy(misValue);
            range2ndInsertionX1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            Excel.Range range2ndInsertion2 = xlWorkSheet.UsedRange.get_Range("BQ1", "BS1") as Excel.Range;
            range2ndInsertion2.EntireColumn.Copy(misValue);
            range2ndInsertionX2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            Excel.Range range2ndInsertion3 = xlWorkSheet.UsedRange.get_Range("BT1", "BV1") as Excel.Range;
            range2ndInsertion3.EntireColumn.Copy(misValue);
            range2ndInsertionX3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);


            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
        //
        //// Supprimer les onglet pour prefaceNP.xls
        //
        private void Supprimeronglet_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;//prefaceNP.xls

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets[1] as Excel.Worksheet;
            Excel.Range range = xlWorkSheet.UsedRange;
            Excel.Window xlWindow = xlApp.ActiveWindow;
            xlWindow.SplitColumn = 0;
            xlWindow.SplitRow = 0;
            Excel.Range rangeinit = range.Cells[1,1] as Excel.Range;
            rangeinit.Select();

            Excel.Worksheet sheetHistMacros = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Macros");
            sheetHistMacros.Delete();
            //Excel.Worksheet sheetTypologieIFRS = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Typologie IFRS");
            //sheetTypologieIFRS.Delete();
            Excel.Worksheet stComboNlist = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Combos et listes à cocher");
            stComboNlist.Delete();
            Excel.Worksheet stHistDiversS = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Divers-s");
            stHistDiversS.Delete();
            ///////////////////////////////////
            //Excel.Worksheet stComposantes = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            //stComposantes.Delete();
            //Excel.Worksheet sheetJ = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            //sheetJ.Delete();
            //Excel.Worksheet stFactgeneraux = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            //stFactgeneraux.Delete();
            //Excel.Worksheet sheetL = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            //sheetL.Delete();
            //Excel.Worksheet sheetM = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            //sheetM.Delete();
            ////////////////////////////////////
            Excel.Worksheet stRappelRet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            stRappelRet.UsedRange.ClearContents();
            //stRappelRet.Delete();
            Excel.Worksheet stTauxParamAr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            stTauxParamAr.UsedRange.ClearContents();
            //stTauxParamAr.Delete();
            Excel.Worksheet stFiscalite = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            stFiscalite.UsedRange.ClearContents();
            //stFiscalite.Delete();
            Excel.Worksheet stFondsDeCommerce = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            stFondsDeCommerce.UsedRange.ClearContents();
            //stFondsDeCommerce.Delete();
            Excel.Worksheet stModuleWacc = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            stModuleWacc.UsedRange.ClearContents();
            //stModuleWacc.Delete();
            Excel.Worksheet stEVAMVA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("EVA-MVA");
            stEVAMVA.Delete();
            Excel.Worksheet stOptionReel = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("OptionsRéelles");
            stOptionReel.Delete();
            Excel.Worksheet stModeleTourT = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModèleTourdetable");
            stModeleTourT.Delete();
            Excel.Worksheet stTourDeTable = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            stTourDeTable.UsedRange.ClearContents();
            //stTourDeTable.Delete();
            Excel.Worksheet stInvestisseur1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investisseur1");
            stInvestisseur1.Delete();
            Excel.Worksheet stInvestisseur2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investisseur2");
            stInvestisseur2.Delete();
            Excel.Worksheet stInvestisseur3 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investisseur3");
            stInvestisseur3.Delete();
            Excel.Worksheet stInvestisseurI = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("InvestisseurI");
            stInvestisseurI.Delete();
            Excel.Worksheet stAddinmenu = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Addin menu");
            stAddinmenu.Delete();
            Excel.Worksheet stMacros = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            stMacros.UsedRange.ClearContents();
            //stMacros.Delete();
            Excel.Worksheet stSensibilite = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Sensibilité");
            stSensibilite.Delete();


            xlApp.Save(misValue);
            xlApp.Quit();
        }
        //
        ////diviser ALL output Histo.xls
        //
        private void Diviser_Click(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            fichierprepare = textBox9.Text;
            prefaceNP = "D:\\ptw\\Histo.xlsx";

            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
         
            
            //petite corr
            object[,] values = (object[,])range.Value2;
            int rCnt = 0;
            int cCnt = 0;
            int row242000 = 0;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row242000 = cf.FindCodedRow("242000-12000", range);

            //cCnt = range.Columns.Count;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "242000-12000"))
            //    {
            //        row242000 = rCnt;
            //        break;
            //    }
            //}
            //Excel.Range cell253F = range.Cells[row242000, 6] as Excel.Range;
            //Excel.Range cell253I = range.Cells[row242000, 9] as Excel.Range;
            //cell253F.Formula = "=C267";
            //cell253I.Formula = "=F267";

            //Hist.Refer coller value
            Excel.Worksheet sheetHistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            Excel.Range rangeHistRefer = sheetHistRefer.UsedRange;
            rangeHistRefer.Copy(misValue);
            rangeHistRefer.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            ////Hist.Refer mise à zero// dégalage?
            object[,] valuesRefer = (object[,])sheetHistRefer.UsedRange.Value2;

            for (int rowCnt = 1; rowCnt <= rangeHistRefer.Rows.Count-1; rowCnt++)//sauf derniere ligne..
            {
                string valuecellabs = Convert.ToString(valuesRefer[rowCnt, 1]);
                if (valuecellabs != "")
                {
                    Excel.Range referZero = sheetHistRefer.UsedRange.get_Range(sheetHistRefer.UsedRange.Cells[rowCnt, 3], sheetHistRefer.UsedRange.Cells[rowCnt, 11]) as Excel.Range;
                    //referZero.Copy();
                    referZero.Value2 = 0;
                    //D1 D2 D3  =""
                    if (valuecellabs == "D" || valuecellabs == "D1" || valuecellabs == "d")
                    {
                        referZero.Formula = "=\"\"";
                    }
                }
            }
            

            //suppression des onglets
            Excel.Worksheet sheetpreface = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet sheetCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-s");
            Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-s");
            Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-s");
            Excel.Worksheet HistReferssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");

            Excel.Worksheet ComptesannuelRefssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Worksheet Comptesannuelssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            Excel.Worksheet Paramimprsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            Excel.Worksheet Paramgenerauxsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            Excel.Worksheet AdminLanguessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            Excel.Worksheet AdminServicesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            Excel.Worksheet Tsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            Excel.Worksheet ParamSavsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            Excel.Worksheet Macrossheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            Excel.Worksheet Vsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            Excel.Worksheet Mosaiquesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            Excel.Worksheet GraphiquesSRsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            Excel.Worksheet Graphimprsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            Excel.Worksheet Dontdeletesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            Excel.Worksheet Finsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            Excel.Worksheet ChoixMethodessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            Excel.Worksheet Noterecapitulativesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            Excel.Worksheet SyntheseValorisationssheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Worksheet DefinitionsArrierePlansheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            Excel.Worksheet RappelRetraitementssheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            Excel.Worksheet RisqueEntreprisesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            Excel.Worksheet ChoixTauxParamsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            Excel.Worksheet TauxParamArrierePlansheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            Excel.Worksheet CorrectifsSIGBilansheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            Excel.Worksheet APNNEsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            Excel.Worksheet FiscaliteDiffereesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            Excel.Worksheet PatrimonialAncAnccsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            Excel.Worksheet FondsDeCommercesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            Excel.Worksheet Goodwillsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            Excel.Worksheet AutresCapitalisationssheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            Excel.Worksheet Multiplessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            Excel.Worksheet MethodesMixtessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            Excel.Worksheet TransactionsComparablessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            Excel.Worksheet GordonShapiroBatessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            Excel.Worksheet CalculFCFsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            Excel.Worksheet DiscountedFCFsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            Excel.Worksheet CmpcWaccsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            Excel.Worksheet CmpcWaccArrierePlansheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            Excel.Worksheet ModuleWaccsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            Excel.Worksheet CCEFsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            Excel.Worksheet TriRentabiliteProjetsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            Excel.Worksheet TourDeTableSynthesesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            Excel.Worksheet EvalLanguessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            Excel.Worksheet Controlessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            Excel.Worksheet EvalServicesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            Excel.Worksheet Composantessheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            Excel.Worksheet Jsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            Excel.Worksheet Factgenerauxsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            Excel.Worksheet Lsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            Excel.Worksheet Msheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            Excel.Worksheet Tresoreriesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            Excel.Worksheet ABsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            Excel.Worksheet Paramtresorsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            Excel.Worksheet Saisonnalitesheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            Excel.Worksheet Zsheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
          //  Excel.Worksheet delete1 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("PreviNotaPme");
            Excel.Worksheet delete2 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("Correctifs.Refer");
          //  delete1.Delete();
            Excel.Worksheet sheetCA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CA");
            Excel.Worksheet sheetInvestissements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investissements");
            Excel.Worksheet sheetCpteresultat = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Cpte Résultat");
            Excel.Worksheet sheetFinancements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Financements");
            Excel.Worksheet sheetbfr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("BFR");
            Excel.Worksheet sheetbilan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Bilan");
            Excel.Worksheet sheetcontrole2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles (2)");
            Excel.Worksheet sheetmultiple = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiple");
            Excel.Worksheet sheetvalo = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Valo et ouverture du capital");
            Excel.Worksheet sheetplan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Plan de financement");
            Excel.Worksheet sheetsynthese = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Synthèse SIG et SR");

            sheetCA.Delete();
            sheetInvestissements.Delete();
            sheetCpteresultat.Delete();
            sheetFinancements.Delete();
            sheetbfr.Delete();
            sheetbilan.Delete();
            sheetcontrole2.Delete();
            sheetmultiple.Delete();
            sheetvalo.Delete();
            sheetplan.Delete();
            sheetsynthese.Delete();

            delete2.Delete();
            model.Delete();
            sheetpreface.Delete();
            sheetCalculs.Delete();
            Historiquesheet.Delete();
            HistPrefacsheet.Delete();
            HistCalculssheet.Delete();
            HistLanguessheet.Delete();
            HistReferssheet.Delete();


            ComptesannuelRefssheet.Delete();
            Comptesannuelssheet.Delete();
            Osheet.Delete();
            Identitesheet.Delete();
            Paramimprsheet.Delete();
            Psheet.Delete();
            Paramgenerauxsheet.Delete();
            AdminLanguessheet.Delete();
            AdminServicesheet.Delete();
            Tsheet.Delete();
            ParamSavsheet.Delete();
            Macrossheet.Delete();
            Vsheet.Delete();
            Mosaiquesheet.Delete();
            GraphiquesSRsheet.Delete();
            Graphimprsheet.Delete();
            Dontdeletesheet.Delete();
            Finsheet.Delete();
            ChoixMethodessheet.Delete();
            Noterecapitulativesheet.Delete();
            SyntheseValorisationssheet.Delete();
            DefinitionsArrierePlansheet.Delete();
            RappelRetraitementssheet.Delete();
            RisqueEntreprisesheet.Delete();
            ChoixTauxParamsheet.Delete();
            TauxParamArrierePlansheet.Delete();
            CorrectifsSIGBilansheet.Delete();
            APNNEsheet.Delete();
            FiscaliteDiffereesheet.Delete();
            PatrimonialAncAnccsheet.Delete();
            FondsDeCommercesheet.Delete();
            Goodwillsheet.Delete();
            AutresCapitalisationssheet.Delete();
            Multiplessheet.Delete();
            MethodesMixtessheet.Delete();
            TransactionsComparablessheet.Delete();
            GordonShapiroBatessheet.Delete();
            CalculFCFsheet.Delete();
            DiscountedFCFsheet.Delete();
            CmpcWaccsheet.Delete();
            CmpcWaccArrierePlansheet.Delete();
            ModuleWaccsheet.Delete();
            CCEFsheet.Delete();
            TriRentabiliteProjetsheet.Delete();
            TourDeTableSynthesesheet.Delete();
            EvalLanguessheet.Delete();
            Controlessheet.Delete();
            EvalServicesheet.Delete();
            Composantessheet.Delete();
            Jsheet.Delete();
            Factgenerauxsheet.Delete();
            Lsheet.Delete();
            Msheet.Delete();
            Tresoreriesheet.Delete();
            ABsheet.Delete();
            Paramtresorsheet.Delete();
            Saisonnalitesheet.Delete();
            Zsheet.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            supprimermoin2_Click(sender, e);
            //subdiviser 9000
            button4_Click(sender, e);

            int time2 = System.Environment.TickCount;
            int times = (time2 - time1) / 1000;
            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;

            //timdiviser = Convert.ToString(Convert.ToDecimal(times) / 1000);
            timdiviser = hours + " heures " + minuit + " minutes " + second;

            //MessageBox.Show("jobs done " + tim + " seconds used");
        }
        //
        ////leger preface.xls pour nota-pme
        //
        private void leger_Click(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            textBox20.AppendText("==> Start Création de PrefaceNP.xlsx"+System.Environment.NewLine);
            try
            {
                fichierprepare = textBox11.Text;
                prefaceNP = "D:\\ptw\\prefaceNP.xlsx";

                supprimerTypologie_Click(sender, e);

                button2_Click(sender, e);
                HistoCalculs();

                HistoMettreZero_Click(sender, e);
                HistoRempl_Click(sender, e);

                HistoAuAvAw_Click(sender, e);
                colCE_Click(sender, e);//72000
                supprimerREF_Click(sender, e);

                ////////////Histo.ptw et histo.preface
                button1_Click(sender, e);//Inserer les colonnes correctifs
                Histopreface_Click(sender, e);

                ////////Annuel .ptw
                AnnuelO_Click(sender, e);
                //ComptesAnnuels_Click(sender, e);

                supprimercol_Click(sender, e);
                button5_Click(sender, e);//supprimer ligne -1

                //supprimer les onglets
                Supprimeronglet_Click(sender, e);

                //traitement REF!
                Historique84000();
                //fonctionRemplacerD1();//D1 formule trop longue
                consigneProteger();

                //Pour Hist-s legement
                insertionHistoS(sender, e);//Inserer les colonnes correctifs
                //HistoprefaceHistoS(sender, e);//inserer les colonnes pour Hist.Preface-s
                supprimercolhistoS(sender, e);
                consigneProtegerHistoS();


                changerNumeroligne();//1000-100 pour historique //Param Sav mettre a "1" ---- OLEDB.net
                copyrangeannewlrefer(sender, e);
                button29_Click(sender, e);
                int time2 = System.Environment.TickCount;
                int times = (time2 - time1) / 1000;

                int hours = times / 3600;
                int minuit = times / 60 - hours * 60;
                int second = times - minuit * 60 - hours * 3600;
                timleger = hours + " heures " + minuit + " minutes " + second;
                //timleger = Convert.ToString(Convert.ToDecimal(times) / 1000);
                textBox20.AppendText("Travail terminé : "+timleger+" s" +System.Environment.NewLine);
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString());
            }
        }

        private void formarthistcalculs()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            // prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");

            xlWorkSheet.UsedRange.get_Range("B1", "B1").EntireColumn.Replace("!$C", "!$E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("B1", "B1").EntireColumn.Replace("!C", "!E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("D1", "D1").EntireColumn.Replace("!$F", "!$H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("F1", "F1").EntireColumn.Replace("!$I", "!$K", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("D1", "D1").EntireColumn.Replace("!F", "!H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet.UsedRange.get_Range("F1", "F1").EntireColumn.Replace("!I", "!K", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        //
        ////Annuels diviser et styler pour nota-pme
        //
        private void AnnuelStyle_Click(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;

            fichierprepare = textBox1.Text;
            prefaceNP = fichierprepare;

            AnnuelO_Click(sender, e);
            AllegerAnnuel_Click(sender, e);
            DiviserComptesAnnuel_Click(sender, e);

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);

            MessageBox.Show("Fusion ptw terminée : " + tim + " secondes");
        }
        //
        ////choiser le fonction de button lancer par rapport le radiobutton
        //
        private void buttonlancer_Click(object sender, EventArgs e)
        {
            leger_Click(sender, e);
            MessageBox.Show("Fusion ptw terminée : " + timleger + " secondes");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox20.AppendText("==> Start Découpage des fichiers"+System.Environment.NewLine);
            try
            {
                int timex1 = System.Environment.TickCount;
                if (checkBox19.Checked == true)
                {
                    Diviser_Click(sender, e);//historique
                }
                if (checkBox21.Checked == true)
                {
                    DiviserHistS(sender, e);//historique-s
                }
                if (checkBox22.Checked == true)
                {
                    diviserAnnuel(sender, e);//comptes Annuels
                }
                if (checkBox23.Checked == true)
                {
                    diviserSynthese(sender, e);//SynthèseValorisations
                }
                button19_Click(sender, e);//language formart
                textBox20.AppendText("Découpage Histo.ptw,    " + timdiviser + " secondes" + Environment.NewLine +
                                "Découpage Histo-s.ptw, " + timdiviserHistoS + " secondes" + Environment.NewLine +
                                "Découpage Annuel.ptw, " + timdiviserAnnuel + " secondes" + Environment.NewLine +
                                "Découpage Eval.ptw, " + timdiviserSynthese + " secondes" + Environment.NewLine);
                MessageBox.Show("Découpage Histo.ptw,    " + timdiviser + " secondes" + Environment.NewLine +
                                "Découpage Histo-s.ptw, " + timdiviserHistoS + " secondes" + Environment.NewLine +
                                "Découpage Annuel.ptw, " + timdiviserAnnuel + " secondes" + Environment.NewLine +
                                "Découpage Eval.ptw, " + timdiviserSynthese + " secondes" + Environment.NewLine
                                );
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
        }
        //app style
        private void button30_Click(object sender, EventArgs e)
        {
            if(textBox2.Text != null)
                stylexml = textBox8.Text;
            else
                MessageBox.Show("Veuillez choisir le fichier des Styles en format XML");
            Xmllire_Click(sender, e);
        }


        #endregion

        #region fusionner *.ptw
        //////////////////////////////////////////////////////////////////////////////////////////
        private void Fussioner_FinalClick(object sender, EventArgs e)
        {
            textBox20.AppendText("==> Start fusion des *.ptw " + System.Environment.NewLine);
           
                int timex1 = System.Environment.TickCount;
                Excel.Application xlApp;
                string pathsouce = textBox4.Text.ToString().Trim();
                string disstion = textBox5.Text.ToString().Trim();
                //Excel.Workbook xlWorkBook0;//preface.xls
                Excel.Workbook xlWorkBook;//Annuel.ptw
                Excel.Workbook xlWorkBook2;//Admin.ptw
                Excel.Workbook xlWorkBook3;//Histo.ptw
                Excel.Workbook xlWorkBook4;//Eval.ptw
                Excel.Workbook xlWorkBook5;//Decis.ptw
                Excel.Workbook xlWorkBook6;//Tres.ptw
                Excel.Workbook xlWorkBook7;//Histo-s.ptw
                
                    object misValue = System.Reflection.Missing.Value;
                    int time1 = System.Environment.TickCount;
                    xlApp = new Excel.ApplicationClass();
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;
                    //xlWorkBook0 = xlApp.Workbooks.Open("D:\\ptw\\preface.xls", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook = xlApp.Workbooks.Open("D:\\ptw\\Annuel.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook2 = xlApp.Workbooks.Open("D:\\ptw\\Admin.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook3 = xlApp.Workbooks.Open("D:\\ptw\\Histo.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook4 = xlApp.Workbooks.Open("D:\\ptw\\Eval.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook5 = xlApp.Workbooks.Open("D:\\ptw\\Decis.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook6 = xlApp.Workbooks.Open("D:\\ptw\\Tres.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    //xlWorkBook7 = xlApp.Workbooks.Open("D:\\ptw\\Histo-s.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook = xlApp.Workbooks.Open(pathsouce + "\\Annuel.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook2 = xlApp.Workbooks.Open(pathsouce + "\\Admin.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook3 = xlApp.Workbooks.Open(pathsouce + "\\Histo.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook4 = xlApp.Workbooks.Open(pathsouce + "\\Eval.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook5 = xlApp.Workbooks.Open(pathsouce + "\\Decis.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook6 = xlApp.Workbooks.Open(pathsouce + "\\Tres.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
                    xlWorkBook7 = xlApp.Workbooks.Open(pathsouce + "\\Histo-s.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

                    try
                    {

                    int nsEval = xlWorkBook4.Sheets.Count;
                    int nsDecis = xlWorkBook5.Sheets.Count;
                    int nsTres = xlWorkBook6.Sheets.Count;
                    Excel.Worksheet xlWorkSheet3 = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item("Historique");
                    //Excel.Range all1 = xlWorkSheet3.UsedRange;
                    //all1.Font.Name = "Cambria";
                    //all1.Font.FontStyle = "Normal";
                    //all1.Font.ColorIndex = 1;
                    releaseObject(xlWorkSheet3);

                    ///////Histo-s xlWorkBook7//////////////////////////////////////////
                    int nsHistos = xlWorkBook7.Sheets.Count;
                    for (int nsx = 1; nsx <= xlWorkBook7.Sheets.Count; nsx++)
                    {
                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook7.Worksheets.get_Item(nsx);
                        admin1.Unprotect(misValue);//pour mosaique
                        admin1.Name = admin1.Name.ToString() + "-s";
                        releaseObject(admin1);
                    }
                    for (int ns = 1; ns <= nsHistos; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook7.Worksheets.get_Item(1);
                        admin1.Unprotect(misValue);//pour mosaique
                        admin1.Move(misValue, histolastsheet);

                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                    }

                    ///////Annuel///////////////////////////////////////// xlWorkBook
                    Excel.Worksheet rangeAnnueldel1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
                    Excel.Worksheet rangeAnnueldel2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Liasse CERFA 30-2970 à 73");
                    rangeAnnueldel1.Delete();
                    rangeAnnueldel2.Delete();
                    int nsAnnuel = xlWorkBook.Sheets.Count;
                    for (int ns = 1; ns < nsAnnuel; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        admin1.Unprotect(misValue);//pour mosaique
                        //Excel.Range all = admin1.UsedRange;
                        //all.Font.Name = "Cambria";
                        //all.Font.FontStyle = "Normal";
                        //all.Font.ColorIndex = 1;
                        admin1.Move(misValue, histolastsheet);

                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                    }

                    ///////Admin///////////////////////////////// xlWorkBook2
                    Excel.Worksheet rangeAdminrenome = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item("Langues");
                    Excel.Worksheet rangeAdminrenome2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item("Service");
                    rangeAdminrenome.Name = "Admin.Langues";
                    rangeAdminrenome2.Name = "Admin.Service";

                    int nsAdmin = xlWorkBook2.Sheets.Count;
                    for (int ns = 1; ns <= nsAdmin; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
                        admin1.Unprotect(misValue);//pour mosaique
                        //Excel.Range all = admin1.UsedRange;
                        //all.Font.Name = "Cambria";
                        //all.Font.FontStyle = "Normal";
                        //all.Font.ColorIndex = 1;
                        admin1.Move(misValue, histolastsheet);

                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                    }

                    ///////Eval/////////////// xlWorkBook4
                    Excel.Worksheet rangeEvalrenome = (Excel.Worksheet)xlWorkBook4.Worksheets.get_Item("Langues");
                    Excel.Worksheet rangeEvalrenome2 = (Excel.Worksheet)xlWorkBook4.Worksheets.get_Item("Service");
                    rangeEvalrenome.Name = "Eval.Langues";
                    rangeEvalrenome2.Name = "Eval.Service";
                    for (int ns = 1; ns <= nsEval; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook4.Worksheets.get_Item(1);
                        admin1.Unprotect(misValue);//pour mosaique

                        admin1.Move(misValue, histolastsheet);

                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                    }

                    ///////Decis xlWorkBook5
                    for (int ns = 1; ns <= nsDecis; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);
                        xlWorkBook3.Sheets.Add(misValue, histolastsheet, misValue, misValue);
                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook5.Worksheets.get_Item(ns);
                        admin1.Unprotect(misValue);//pour mosaique
                        Excel.Worksheet admin1X = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        admin1X.Name = admin1.Name.ToString();
                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                        releaseObject(admin1X);
                    }

                    ///////Tres xlWorkBook6
                    for (int ns = 1; ns <= nsTres; ns++)
                    {
                        Excel.Worksheet histolastsheet = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);
                        xlWorkBook3.Sheets.Add(misValue, histolastsheet, misValue, misValue);
                        Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook6.Worksheets.get_Item(ns);
                        admin1.Unprotect(misValue);//pour mosaique
                        Excel.Worksheet admin1X = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(xlWorkBook3.Sheets.Count);

                        admin1X.Name = admin1.Name.ToString();
                        releaseObject(histolastsheet);
                        releaseObject(admin1);
                        releaseObject(admin1X);
                    }


                    xlWorkSheet3 = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item("Historique");
                    Excel.Range adminreplacefinal = xlWorkSheet3.UsedRange;
                    adminreplacefinal.Cells.Replace("[admin.ptw]Langues", "Admin.Langues", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                    adminreplacefinal.Cells.Replace("[admin.ptw]Service", "Admin.Service", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                    Excel.Range fomarts = (Excel.Range)xlWorkSheet3.get_Range("B1", "B" + (adminreplacefinal.Rows.Count - 1));
                    fomarts.Cells.NumberFormat = null;
                    for (int nss = 1; nss <= xlWorkBook3.Sheets.Count; nss++)
                    {
                        Excel.Worksheet SheetRem = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item(nss);
                        Excel.Range rangeRem = SheetRem.UsedRange;

                        rangeRem.Cells.Replace("[Decis.ptw]", "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeRem.Cells.Replace("[Tres.ptw]", "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeRem.Cells.Replace("'D:\\ptw\\Annuel.ptw'!", "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeRem.Cells.Replace("'D:\\ptw\\Admin.ptw'!", "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        releaseObject(rangeRem);
                        releaseObject(SheetRem);
                    }


                    Excel.Worksheet xlWorkSheetParamSav = (Excel.Worksheet)xlWorkBook3.Worksheets.get_Item("Param Sav");

                    Excel.Range range = xlWorkSheetParamSav.UsedRange;
                    object[,] values = (object[,])range.Value2;

                    int rCnt = 0;
                    int cCnt = 0;
                    int colx = 0;
                    int rowx = 0;

                    rCnt = range.Rows.Count;
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                        if (Regex.Equals(valuecellabs, "3000"))
                        {
                            colx = cCnt;
                            break;
                        }
                    }


                    cCnt = range.Columns.Count;
                    for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                    {
                        string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                        if (Regex.Equals(valuecellabs, "267000"))
                        {
                            rowx = rCnt;
                            break;
                        }
                    }

                    Excel.Range rangecellx = xlWorkSheetParamSav.Cells[rowx, colx] as Excel.Range;
                    rangecellx.Value2 = 1;


                    xlWorkSheet3.SaveAs(disstion + "\\preface.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


                    //xlWorkBook3.Close(true, misValue, misValue);
                    xlApp.Quit();


                    int time2 = System.Environment.TickCount;

                    int times = (time2 - time1) / 1000;
                    int hours = times / 3600;
                    int minuit = times / 60 - hours * 60;
                    int second = times - minuit * 60 - hours * 3600;
                    timfusion = hours + " heure(s) " + minuit + " minutes " + second;
                    //timfusion = Convert.ToString(Convert.ToDecimal(times) / 1000);
                    //MessageBox.Show("jobs done " + tim + " seconds used");
                    textBox20.AppendText("Travail terminé " + timfusion + " s" + System.Environment.NewLine);
                    releaseObject(adminreplacefinal);
                    releaseObject(xlWorkSheet3);
                }
                catch (Exception ex)
                {
                    textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
                }
                finally
                {
                    
                    releaseObject(xlWorkBook7);
                    releaseObject(xlWorkBook6);
                    releaseObject(xlWorkBook5);
                    releaseObject(xlWorkBook4);
                    releaseObject(xlWorkBook3);
                    releaseObject(xlWorkBook2);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
        }
        //////////////////////////////////////////////////////////////////////////////////////////       
        #endregion

        #region dégroupage
        //
        //// subdiviser 1 9000
        //
        private void button4_Click(object sender, EventArgs e)
        {
            pathnotapme = textBox3.Text;
            pathstylerfinal = textBox6.Text;

            string openfilex = "D:\\ptw\\Histo.xlsx";

            ////////////////open excel///////////////////////////////////////
            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Historique");
            Excel.Range rangeRefall = xlWorkSheetRef.UsedRange;
            //bug : le seul moyen pour supprimer la dernière colonne est de chnager la largeur de toutes les colonnes (on ne sait pas pourquoi) !!!
            xlWorkSheetRef.Cells.ColumnWidth = 20;

            Excel.Range rangeRef = xlWorkSheetRef.Cells[rangeRefall.Rows.Count, 1] as Excel.Range;
            rangeRef.EntireRow.Copy(misValue);
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", xlWorkSheetRef.Cells[rangeRefall.Rows.Count - 1, 1]) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            xlWorkSheetRef.SaveAs("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            //////////////////////////////////////////////////////////////////////////////////
            Thread.Sleep(3000);
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlApp.Application.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 ="[" + openfilex.Substring(7, 9) + "]";
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            int col83000 = 0;
            rCnt = range.Rows.Count;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col3000 = cf.FindCodedColumn("3000", range);
            col4000 = cf.FindCodedColumn("4000", range);
            col5000 = cf.FindCodedColumn("5000", range);
            col8000 = cf.FindCodedColumn("8000", range);
            col = cf.FindCodedColumn("9000", range);
            col83000 = cf.FindCodedColumn("83000", range);


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "3000"))
            //    {
            //        col3000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "4000"))
            //    {
            //        col4000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "5000"))
            //    {
            //        col5000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "8000"))
            //    {
            //        col8000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "9000"))
            //    {
            //        col = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "83000"))
            //    {
            //        col83000 = cCnt;
            //        break;
            //    }
            //}
            int fileflag = 0;
            for (int row = 25; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    Thread.Sleep(3000);
                    xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Historique");
                   // string[] namestable = { "ACT1.xlsx", "ACT2.xlsx", "ACT3.xlsx", "ACT4.xlsx", "PAS1.xlsx", "PAS2.xlsx", "PAS3.xlsx", "CR1.xlsx", "CR2.xlsx", "CR3.xlsx", "CR4.xlsx", "ANN5-1.xlsx", "ANN5-2.xlsx", "ANN5-3.xlsx", "ANN6-1.xlsx", "ANN6-2.xlsx", "ANN6-3.xlsx", "ANN7-1.xlsx", "ANN7-2.xlsx", "ANN7-3.xlsx", "ANN8-1.xlsx", "ANN8-2.xlsx", "ANN11-1.xlsx" };
                    string[] namestable = { "ACT1.xlsx", "ACT4.xlsx", "PAS1.xlsx", "PAS3.xlsx", "CR1.xlsx", "CR3.xlsx",  "ANN5-1.xlsx", "ANN5-2.xlsx",  "ANN6-1.xlsx", "ANN6-2.xlsx", "ANN6-3.xlsx", "ANN7-1.xlsx", "ANN7-2.xlsx", "ANN7-3.xlsx", "ANN8-1.xlsx", "ANN8-2.xlsx", "ANN11-1.xlsx" };
                    string divisavenom = pathnotapme + "\\" + namestable[fileflag]; 
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire

                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[25, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[25, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[24, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[25, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row --;// point important, pour garder l'ordre de ligne ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range("A25", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;


                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Historique");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);

                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;



                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Historique");
                    Excel.Range rangeremplace = xlWorkSheetremplace.UsedRange;
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas
                    ////////delete col8000 "-2"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    {
                        string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                        if (Regex.Equals(valuedel, "-2"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowdel--;
                        }
                    }
                    ///////////////row hide "-5"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }
                    ///////////////Hide -1 pour col 83000/////////////////////////////////////////////
                    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                    //    if (Regex.Equals(valuedel, "-1"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col83000] as Excel.Range;
                    //        rangeDely.EntireRow.Hidden = true;
                    //    }
                    //}
                    /////////////////////////////////////////////////////////////////////////////////
                    object[,] valuesNX = (object[,])rangeremplace.Value2;
                    //string valueNX = Convert.ToString(valuesNX[row, col]);
                    for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row3000, col8000]) != "-7")//-7 non zero
                        {
                            if (valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "224000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "242000-12000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "275000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "762000-4000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-2000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-4000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "299000-200" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "468000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "471000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "473000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "475000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "478000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "480000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "482000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "485000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "487000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-2000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-3000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-2000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-3000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "768000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "772000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "776000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "780000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-500" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-700" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "814000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "816000-1000" || valuesNX[row3000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "818000-1000" || values[row3000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "308000" || valuesNX[row3000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "347000" || valuesNX[row3000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "353000" || valuesNX[row3000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "350000" || values[row3000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "344000")
                            {
                            }
                            else
                            {

                                rangeDely.Value2 = 0;
                            }
                        }
                    }
                    for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row4000, col8000]) != "-7")//-7 non zero
                        {
                            if (valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "224000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "242000-12000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "275000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "762000-4000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-2000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-4000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "299000-200" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "468000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "471000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "473000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "475000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "478000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "480000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "482000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "485000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "487000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-2000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-3000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-2000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-3000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "768000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "772000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "776000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "780000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-500" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-700" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "814000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "816000-1000" || valuesNX[row4000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "818000-1000" || values[row4000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "308000" || valuesNX[row4000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "347000" || valuesNX[row4000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "353000" || valuesNX[row4000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "350000" || values[row4000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "344000")
                            {
                            }
                            else
                            {
                                rangeDely.Value2 = 0;
                            }
                        }
                    }
                    for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row5000, col8000]) != "-7")//-7 non zero
                        {
                            if (valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "224000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "242000-12000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "275000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "762000-4000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-2000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "763000-4000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "243000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "299000-200" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "468000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "471000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "473000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "475000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "478000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "480000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "482000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "485000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "487000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-2000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "745000-3000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-2000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "746000-3000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "768000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "772000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "776000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "780000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-500" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-700" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "791000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "814000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "816000-1000" || valuesNX[row5000, xlWorkSheetremplace.UsedRange.Columns.Count].ToString() == "818000-1000" || values[row5000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "308000" || valuesNX[row5000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "347000" || valuesNX[row5000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "353000" || valuesNX[row5000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "350000" || values[row5000, xlWorkSheet.UsedRange.Columns.Count].ToString() == "344000")
                            {
                            }
                            else
                            {
                                rangeDely.Value2 = 0;
                            }
                        }
                    }
                    if (namestable[fileflag] == "PAS3.xlsx")
                    {
                        Excel.Worksheet xlWorkSheetremplace2 = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Historique");
                        Excel.Range rangeremplace2 = xlWorkSheetremplace.UsedRange;
                        rangeremplace2.Cells.Replace("Hist.Refer!I20", "Hist.Refer!I20*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!C20", "Hist.Refer!C20*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!I19", "Hist.Refer!I19*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!I18", "Hist.Refer!I18*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!I14", "Hist.Refer!I14*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!I13", "Hist.Refer!I13*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                        rangeremplace2.Cells.Replace("Hist.Refer!F20", "Hist.Refer!F20*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!F19", "Hist.Refer!F19*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!F18", "Hist.Refer!F18*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!F14", "Hist.Refer!F14*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!F13", "Hist.Refer!F13*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

                        rangeremplace2.Cells.Replace("Hist.Refer!C20", "Hist.Refer!C20*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!C19", "Hist.Refer!C19*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!C18", "Hist.Refer!C18*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!C14", "Hist.Refer!C14*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                        rangeremplace2.Cells.Replace("Hist.Refer!C13", "Hist.Refer!C13*1", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                    }
                    ////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    if (checkBox20.Checked == true)
                    {
                        fileAstyler = divisavenom;
                        Xmllire_Click(sender, e);
                    }

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = 25;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// subdiviser Comptes Annuel ------ Annuel.ptw
        //
        private void DiviserComptesAnnuel_Click(object sender, EventArgs e)
        {
            pathnotapme = textBox3.Text;
            pathstylerfinal = textBox6.Text;


            ////////////////open excel///////////////////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;




            xlWorkBookRef = xlAppRef.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Comptes annuels");
            Excel.Range rangeRefall = xlWorkSheetRef.UsedRange;
            //exception!!!
            xlWorkSheetRef.Cells.ColumnWidth = 20;

            Excel.Range rangeRef = xlWorkSheetRef.Cells[rangeRefall.Rows.Count, 1] as Excel.Range;
            rangeRef.EntireRow.Copy(misValue);
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", xlWorkSheetRef.Cells[rangeRefall.Rows.Count - 1, 1]) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            xlWorkSheetRef.SaveAs("D:\\ptw\\Annuelx.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();

            releaseObject(xlWorkSheetRef);
            releaseObject(xlWorkBookRef);
            releaseObject(xlAppRef);
            //////////////////////////////////////////////////////////////////////////////////
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls

            string remplacehisto8 = "[Annuel.ptw]";
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            //int col8000 = 0;
            int col11000 = 0;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col3000 = cf.FindCodedColumn("3000", range);
            col4000 = cf.FindCodedColumn("4000", range);
            col5000 = cf.FindCodedColumn("5000", range);
            col = cf.FindCodedColumn("8000", range);
            col11000 = cf.FindCodedColumn("11000-1000", range);

            rCnt = range.Rows.Count;


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "3000"))
            //    {
            //        col3000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "4000"))
            //    {
            //        col4000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "5000"))
            //    {
            //        col5000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "8000"))//rupture ref 8000
            //    {
            //        col = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "11000-1000"))
            //    {
            //        col11000 = cCnt;
            //        break;
            //    }
            //}

            int fileflag = 0;
            for (int row = 25; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Annuelx.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Comptes annuels");
                    string[] namestable = { "CR1.xlsx", "CR2.xlsx", "CR3.xlsx", "Actif1.xlsx", "Passif1.xlsx", "Flux1.xlsx", "Flux2.xlsx", "Ratio1.xlsx", "Ratio2.xlsx", "SynthExpl.xlsx", "SynthBilan.xlsx", "SynthStructure.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire
                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[19, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[19, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[18, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[19, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de row ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range("A19", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Comptes annuels");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);
                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBooknewx1.Close(true, misValue, misValue);







                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Comptes annuels");
                    Excel.Range rangeremplace = xlWorkSheetremplace.UsedRange;
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas









                    ////////delete col11000 "-6"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    //for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowdel, col11000]);
                    //    if (Regex.Equals(valuedel, "-6"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col11000] as Excel.Range;
                    //        rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    //        rangeremplace = xlWorkSheetremplace.UsedRange;
                    //        values8000 = (object[,])rangeremplace.Value2;
                    //        rowdel--;
                    //    }
                    //}



                    ///////////////row hide "-5"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col11000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col11000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }


                    /////////////////////////////////////////////////////////////////////////////
                    //object[,] valuesNX = (object[,])rangeremplace.Value2;
                    ////string valueNX = Convert.ToString(valuesNX[row, col]);
                    //for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True")
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}
                    //for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True")
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}
                    //for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True")
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}
                    //////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    if (checkBox1.Checked == true)
                    {
                        fileAstyler = divisavenom;
                        XmllireAnnuel_Click(sender, e);
                    }

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = 19;//important remise le ligne commencer apres action delete
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        //// Appliquer Style bloc par bloc, fonction test, a supprimer
        //
        private void Astyle_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open("D:\\ptw\\Histo6.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");

            Excel.Range range = xlWorkSheet.UsedRange;

            Excel.Range rangex1 = xlWorkSheet.UsedRange.get_Range("A21", "E21") as Excel.Range;
            Excel.Range rangex2 = xlWorkSheet.UsedRange.get_Range("A24", "E24") as Excel.Range;
            Excel.Range rangex3 = xlWorkSheet.UsedRange.get_Range("A26", "E26") as Excel.Range;
            Excel.Range rangex4 = xlWorkSheet.UsedRange.get_Range("A27", "E27") as Excel.Range;


            Excel.Range rangestyle = xlWorkSheet.UsedRange.get_Range("A30", "E40") as Excel.Range;
            rangex1.Copy(misValue);
            rangestyle.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);

            //xlApp.CutCopyMode = Microsoft.Office.Interop.Excel.XlCutCopyMode.none;

            //----------------------------------------------------------------------------------------
            //Excel.Style style1 = xlWorkBook.Styles.Add("NewStyle", misValue);
            //style1.Font.Name = "Verdana";
            //style1.Font.Size = 12;
            //style1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            //style1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            //style1.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            ////MessageBox.Show(rangex1.Style.ToString());
            //rangestyle.Style = "NewStyle";
            //----------------------------------------------------------------------------------------
        }
        #endregion

        //flush all Excel object dans memoire
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        #region test ini fonction test performance de fichier ini, a supprimer
        ////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////
        ////Test appliquer style Methode copie coller
        //
        private void button13_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open("D:\\Histo test.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount; 
            int col = 10;
            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1"))
                {
                    Excel.Range rangeStyleSource = xlWorkSheet.UsedRange.get_Range("A1", "F1") as Excel.Range;
                    rangeStyleSource.Copy(misValue);
                    Excel.Range rangeTerminal = xlWorkSheet.UsedRange.get_Range("A" + row, "D" + row) as Excel.Range;
                    rangeTerminal.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                }
                if (Regex.Equals(value, "2"))
                {
                    Excel.Range rangeStyleSource = xlWorkSheet.UsedRange.get_Range("A2", "F2") as Excel.Range;
                    rangeStyleSource.Copy(misValue);
                    Excel.Range rangeTerminal = xlWorkSheet.UsedRange.get_Range("A" + row, "D" + row) as Excel.Range;
                    rangeTerminal.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                }
                if (Regex.Equals(value, "3"))
                {
                    Excel.Range rangeStyleSource = xlWorkSheet.UsedRange.get_Range("A3", "F3") as Excel.Range;
                    rangeStyleSource.Copy(misValue);
                    Excel.Range rangeTerminal = xlWorkSheet.UsedRange.get_Range("A" + row, "D" + row) as Excel.Range;
                    //Excel.Range rangeTerminal = xlWorkSheet.Cells[row, 1] as Excel.Range;
                    rangeTerminal.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                }
            }

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);

            MessageBox.Show(tim + "secondes");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //
        ////Test appliquer style Methode cellule par cellule
        //
        private void button14_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open("D:\\Histo test.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //----------------------------------------------------------------------------------------
            Excel.Style style11 = xlWorkBook.Styles.Add("NewStyle11", misValue);
            style11.Font.Name = "Verdana";
            style11.Font.Size = 12;
            style11.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            style11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            style11.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //----------------------------------------------------------------------------------------
            //----------------------------------------------------------------------------------------
            Excel.Style style22 = xlWorkBook.Styles.Add("NewStyle22", misValue);
            style22.Font.Name = "Ariel";
            style22.Font.Size = 10;
            style22.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumTurquoise);
            style22.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gainsboro);
            style22.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //----------------------------------------------------------------------------------------
            //----------------------------------------------------------------------------------------
            Excel.Style style33 = xlWorkBook.Styles.Add("NewStyle33", misValue);
            style33.Font.Name = "Verdana";
            style33.Font.Size = 8;
            style33.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            style33.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            style33.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //----------------------------------------------------------------------------------------
            //----------------------------------------------------------------------------------------
            Excel.Style style44 = xlWorkBook.Styles.Add("NewStyle44", misValue);
            style44.Font.Name = "Verdana";
            style44.Font.Size = 6;
            style44.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
            style44.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
            style44.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            //----------------------------------------------------------------------------------------
            int time1=System.Environment.TickCount; 
            int col = 10;
            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1"))
                {
                    Excel.Range rangeTerminal = xlWorkSheet.Cells[row, 1] as Excel.Range;
                    rangeTerminal.Style = style11;
                    Excel.Range rangeTerminal2 = xlWorkSheet.Cells[row, 2] as Excel.Range;
                    rangeTerminal2.Style = style22;
                    Excel.Range rangeTerminal3 = xlWorkSheet.Cells[row, 3] as Excel.Range;
                    rangeTerminal3.Style = style33;
                    Excel.Range rangeTerminal4 = xlWorkSheet.Cells[row, 4] as Excel.Range;
                    rangeTerminal4.Style = style44;
                }
                if (Regex.Equals(value, "2"))
                {
                    Excel.Range rangeTerminal = xlWorkSheet.Cells[row, 1] as Excel.Range;
                    rangeTerminal.Style = style44;
                    Excel.Range rangeTerminal2 = xlWorkSheet.Cells[row, 2] as Excel.Range;
                    rangeTerminal2.Style = style33;
                    Excel.Range rangeTerminal3 = xlWorkSheet.Cells[row, 3] as Excel.Range;
                    rangeTerminal3.Style = style22;
                    Excel.Range rangeTerminal4 = xlWorkSheet.Cells[row, 4] as Excel.Range;
                    rangeTerminal4.Style = style11;
                }
                if (Regex.Equals(value, "3"))
                {
                    Excel.Range rangeTerminal = xlWorkSheet.Cells[row, 1] as Excel.Range;
                    rangeTerminal.Style = style22;
                    Excel.Range rangeTerminal2 = xlWorkSheet.Cells[row, 2] as Excel.Range;
                    rangeTerminal2.Style = style44;
                    Excel.Range rangeTerminal3 = xlWorkSheet.Cells[row, 3] as Excel.Range;
                    rangeTerminal3.Style = style11;
                    Excel.Range rangeTerminal4 = xlWorkSheet.Cells[row, 4] as Excel.Range;
                    rangeTerminal4.Style = style33;
                }
            }
            int time2=System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);

            MessageBox.Show(tim+"secondes");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        //ecrire ini
        private void ecrire_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.InitialDirectory = "D:\\ptw\\appstyle.xlsx";
            OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            //OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;
            OpenFileDialog1.ShowDialog();


            ////////////////open excel////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            string openfilex = OpenFileDialog1.FileName.ToString();//"D:\\ptw\\appstyle.xls"
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            string filePath = "D:\\demo.ini";
            IniFile iniFile = new IniFile(filePath);

            int col = 11;
            int time1 = System.Environment.TickCount;

            for (int row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, row.ToString()))
                {
                    string colcount = iniFile.ReadInivalue("style"+row , "col");
                    //MessageBox.Show(colcount);
                    int colcountx = Convert.ToInt32(colcount);
                    for (int colc = 1; colc <= colcountx; colc++)
                    {
                        Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                        string fontname = rangeDelx.Font.Name.ToString();
                        string fontsize = rangeDelx.Font.Size.ToString();
                        string fontcolor = rangeDelx.Font.Color.ToString();
                        string fontstyle = rangeDelx.Font.FontStyle.ToString();
                        string fontbold = rangeDelx.Font.Bold.ToString();
                        string fontitalic = rangeDelx.Font.Italic.ToString();
                        string fontunderline = rangeDelx.Font.Underline.ToString();
                        string bgcolor = rangeDelx.Interior.Color.ToString();
                        string bgcolorindex = rangeDelx.Interior.ColorIndex.ToString();
                        string bordertop = rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle.ToString();
                        string borderbot = rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle.ToString();
                        string borderleft = rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle.ToString();
                        string borderright = rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle.ToString();
                        iniFile.WriteInivalue("style" + row, "font" + colc, fontname);
                        iniFile.WriteInivalue("style" + row, "fontsize" + colc, fontsize);
                        iniFile.WriteInivalue("style" + row, "fontcolor" + colc, fontcolor);
                        iniFile.WriteInivalue("style" + row, "fontstyle" + colc, fontstyle);
                        iniFile.WriteInivalue("style" + row, "fontbold" + colc, fontbold);
                        iniFile.WriteInivalue("style" + row, "fontitalic" + colc, fontitalic);
                        iniFile.WriteInivalue("style" + row, "fontunderline" + colc, fontunderline);
                        iniFile.WriteInivalue("style" + row, "bgcolor" + colc, bgcolor);
                        iniFile.WriteInivalue("style" + row, "bgcolorindex" + colc, bgcolorindex);
                        iniFile.WriteInivalue("style" + row, "bordertop" + colc, bordertop);
                        iniFile.WriteInivalue("style" + row, "borderbot" + colc, borderbot);
                        iniFile.WriteInivalue("style" + row, "borderleft" + colc, borderleft);
                        iniFile.WriteInivalue("style" + row, "borderright" + colc, borderright);

                    }
                    
                }
            }

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            MessageBox.Show("Fusion ptw terminée : " + tim + " secondes");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            //string[] fileNames = new string[] { "font", "fontsize", "colorbg", "colorfont" };
            //string[] values = new string[] { "value1", "value2", "value3", "value4" };
            //for (int i = 0; i < 4; i++)
            //{
            //    iniFile.WriteInivalue("style3", fileNames[i], values[i]);
            //}
        }
        //lire ini
        private void lire_Click(object sender, EventArgs e)
        {
            string filePath = "D:\\demo.ini";
            IniFile iniFile = new IniFile(filePath);


            string singleValue1 = null;
            string singleValue2 = null;
            string singleValue3 = null;
            string singleValue4 = null;
            singleValue1 = iniFile.ReadInivalue("style1", "font");
            singleValue2 = iniFile.ReadInivalue("style1", "fontsize");
            singleValue3 = iniFile.ReadInivalue("style1", "colorbg");
            singleValue4 = iniFile.ReadInivalue("style1", "colorfont");
            MessageBox.Show(singleValue1 + "   " + singleValue2 + "   " + singleValue3 + "   " + singleValue4);


            //string[] fileNames = new string[] { "file1", "file2", "file3", "file4" };
            //ArrayList values = new ArrayList();
            //for (int i = 0; i < 4; i++)
            //{
            //    values.Add(iniFile.ReadInivalue("style3", fileNames[i]));
            //}
            //int nCount = values.Count;
            //string multiValues = null;
            //for (int i = 0; i < nCount; i++)
            //{
            //    multiValues += values[i].ToString() + " ";
            //}
            //MessageBox.Show(multiValues);
        }
        /////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////////////////////////////////////////////////////////
        #endregion

        #region création de fichier style XML
        //
        ////Xml Ecrire
        //
        private void Xmlecrire_Click(object sender, EventArgs e)
        {
            //OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.InitialDirectory = "D:\\ptw\\appstyle.xls";
            //OpenFileDialog1.Filter = "Excel Files .xls|*.xls|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            ////OpenFileDialog1.FilterIndex = 2;
            ////OpenFileDialog1.RestoreDirectory = true;
            //OpenFileDialog1.ShowDialog();
            try
            {
                textBox20.AppendText("==> Start Création du fichier xml des styles" + System.Environment.NewLine);
                ////////////////open excel////////////////////////
                int timet = System.Environment.TickCount;
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet = null;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;

                //string openfilex = OpenFileDialog1.FileName.ToString();//"D:\\ptw\\appstyle.xls"
                string openfilex = textBox7.Text.ToString();
                xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                // xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


                XmlDocument appstyleDoc = new XmlDocument();
                XmlElement styleRoot;
                XmlNode stylexmlNode;

                stylexmlNode = appstyleDoc.CreateNode(XmlNodeType.XmlDeclaration, "", "");
                appstyleDoc.AppendChild(stylexmlNode);
                styleRoot = appstyleDoc.CreateElement("style");
                appstyleDoc.AppendChild(styleRoot);


                ///////////////predifinie palette index////////////////////////////////////////////
                XmlNode racinestyle = appstyleDoc.SelectSingleNode("//style");
                XmlElement xpalette = appstyleDoc.CreateElement("palette");
                racinestyle.AppendChild(xpalette);

                XmlElement nbstyle = appstyleDoc.CreateElement("nbstyle");
                racinestyle.AppendChild(nbstyle);

                for (int nindex = 1; nindex <= 56; nindex++)
                {
                    XmlElement indexx = appstyleDoc.CreateElement("index" + nindex);
                    xpalette.AppendChild(indexx);
                    indexx.InnerText = xlWorkBook.get_Colors(nindex).ToString();
                    ////////////////////////////////////////////attribut RGB faculté
                    int colorindexox = Convert.ToInt32(indexx.InnerText);
                    int colorindexB = colorindexox / 65536;
                    int colorindexG = (colorindexox % 65536) / 256;
                    int colorindexR = (colorindexox % 65536) % 256;
                    indexx.SetAttribute("R", colorindexR.ToString());
                    indexx.SetAttribute("G", colorindexG.ToString());
                    indexx.SetAttribute("B", colorindexB.ToString());
                }
                ///////////////////////////////////////////////////////////////////////////////////

                //routine pour multiple onglet
                //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
                int wkbookcount = xlWorkBook.Worksheets.Count;
                int nb = 1;
                int time1 = System.Environment.TickCount;
                for (int k = 1; k <= wkbookcount-1; k++)
                {

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(k);

                    Excel.Range range = xlWorkSheet.UsedRange;
                    object[,] values = (object[,])range.Value2;



                    int rCnt = 0;
                    int cCnt = 0;
                    int col = 0;
                    rCnt = range.Rows.Count;
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                        if (Regex.Equals(valuecellabs, "1000"))
                        {
                            col = cCnt;
                            break;
                        }
                    }
                    string colcount = "";


                    for (int row = 1; row <= values.GetUpperBound(0) - 1; row++)
                    {
                        string value = Convert.ToString(values[row, col]);
                        if (value != "")
                        {
                            XmlElement numerostyle = appstyleDoc.CreateElement("nbstyle" + nb);
                            nb++;
                            nbstyle.AppendChild(numerostyle);
                            numerostyle.InnerText = value;
                            ////////////////////////////////////////////////////////////////////
                            XmlElement styleNX = appstyleDoc.CreateElement("style" + value);
                            racinestyle.AppendChild(styleNX);
                            XmlElement colNX = appstyleDoc.CreateElement("col");
                            colNX.InnerText = Convert.ToString(values[row, col + 1]);
                            styleNX.AppendChild(colNX);


                            XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + value);
                            if (xstyle != null)
                            {
                                colcount = (xstyle.SelectSingleNode("col")).InnerText;
                            }

                            //MessageBox.Show(colcount);
                            int colcountx = Convert.ToInt32(colcount);
                            for (int colc = 1; colc <= colcountx; colc++)
                            {
                                XmlElement nodeN = appstyleDoc.CreateElement("style" + value + "." + colc);
                                xstyle.AppendChild(nodeN);

                                Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc + 2] as Excel.Range;
                                string fontname = rangeDelx.Font.Name.ToString();
                                string fontsize = rangeDelx.Font.Size.ToString();
                                string fontcolor = rangeDelx.Font.Color.ToString();
                                int fontcnumber = Convert.ToInt32(fontcolor);
                                int colorB = fontcnumber / 65536;
                                int colorG = (fontcnumber % 65536) / 256;
                                int colorR = (fontcnumber % 65536) % 256;
                                string fontcolorindex = rangeDelx.Font.ColorIndex.ToString();
                                string fontbold = rangeDelx.Font.Bold.ToString();
                                string fontitalic = rangeDelx.Font.Italic.ToString();
                                string fontunderline = rangeDelx.Font.Underline.ToString();
                                string bgcolor = rangeDelx.Interior.Color.ToString();
                                int bgcnumber = Convert.ToInt32(bgcolor);
                                int bgcolorB = bgcnumber / 65536;
                                int bgcolorG = (bgcnumber % 65536) / 256;
                                int bgcolorR = (bgcnumber % 65536) % 256;

                                string bgcolorindex = rangeDelx.Interior.ColorIndex.ToString();
                                //top
                                string bordertop = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle.ToString();
                                string borderweighttop = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight.ToString();
                                //bot
                                string borderbot = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.ToString();
                                string borderweightbot = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight.ToString();
                                //left
                                string borderleft = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle.ToString();
                                string borderweightleft = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight.ToString();
                                //right
                                string borderright = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle.ToString();
                                string borderweightright = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight.ToString();

                                string wraptext = rangeDelx.WrapText.ToString();
                                string Halignment = rangeDelx.HorizontalAlignment.ToString();
                                string Valignment = rangeDelx.VerticalAlignment.ToString();
                                string mergecell = rangeDelx.MergeCells.ToString();
                                string mergecellcount = rangeDelx.MergeArea.Cells.Count.ToString();

                                string nomberformat = rangeDelx.NumberFormat.ToString();

                                string locked = rangeDelx.Locked.ToString();
                                string formulahidden = rangeDelx.FormulaHidden.ToString();

                                string colwidth = rangeDelx.ColumnWidth.ToString();
                                string rowheight = rangeDelx.RowHeight.ToString();

                                //
                                XmlElement nodeN1 = appstyleDoc.CreateElement("font");
                                nodeN.AppendChild(nodeN1);
                                nodeN1.InnerText = fontname;
                                //
                                XmlElement nodeN2 = appstyleDoc.CreateElement("fontsize");
                                nodeN.AppendChild(nodeN2);
                                nodeN2.InnerText = fontsize;
                                //
                                XmlElement nodeN3 = appstyleDoc.CreateElement("fontcolorindex");
                                nodeN.AppendChild(nodeN3);
                                nodeN3.InnerText = fontcolorindex;
                                //
                                XmlElement nodeN3a = appstyleDoc.CreateElement("fontcolor");
                                nodeN.AppendChild(nodeN3a);
                                nodeN3a.InnerText = fontcolor;
                                nodeN3a.SetAttribute("R", colorR.ToString());
                                nodeN3a.SetAttribute("G", colorG.ToString());
                                nodeN3a.SetAttribute("B", colorB.ToString());
                                //
                                XmlElement nodeN5 = appstyleDoc.CreateElement("fontbold");
                                nodeN.AppendChild(nodeN5);
                                nodeN5.InnerText = fontbold;
                                //
                                XmlElement nodeN6 = appstyleDoc.CreateElement("fontitalic");
                                nodeN.AppendChild(nodeN6);
                                nodeN6.InnerText = fontitalic;
                                //
                                XmlElement nodeN7 = appstyleDoc.CreateElement("fontunderline");
                                nodeN.AppendChild(nodeN7);
                                nodeN7.InnerText = fontunderline;
                                //
                                XmlElement nodeN8 = appstyleDoc.CreateElement("bgcolor");
                                nodeN.AppendChild(nodeN8);
                                nodeN8.InnerText = bgcolor;
                                nodeN8.SetAttribute("R", bgcolorR.ToString());
                                nodeN8.SetAttribute("G", bgcolorG.ToString());
                                nodeN8.SetAttribute("B", bgcolorB.ToString());
                                //
                                XmlElement nodeN8a = appstyleDoc.CreateElement("bgcolorindex");
                                nodeN.AppendChild(nodeN8a);
                                nodeN8a.InnerText = bgcolorindex;
                                //
                                XmlElement nodeN9 = appstyleDoc.CreateElement("bordertop");
                                nodeN.AppendChild(nodeN9);
                                nodeN9.InnerText = bordertop;
                                XmlElement nodeN9a = appstyleDoc.CreateElement("borderweighttop");
                                nodeN.AppendChild(nodeN9a);
                                nodeN9a.InnerText = borderweighttop;
                                //
                                XmlElement nodeN10 = appstyleDoc.CreateElement("borderbot");
                                nodeN.AppendChild(nodeN10);
                                nodeN10.InnerText = borderbot;
                                XmlElement nodeN10a = appstyleDoc.CreateElement("borderweightbot");
                                nodeN.AppendChild(nodeN10a);
                                nodeN10a.InnerText = borderweightbot;
                                //
                                XmlElement nodeN11 = appstyleDoc.CreateElement("borderleft");
                                nodeN.AppendChild(nodeN11);
                                nodeN11.InnerText = borderleft;
                                XmlElement nodeN11a = appstyleDoc.CreateElement("borderweightleft");
                                nodeN.AppendChild(nodeN11a);
                                nodeN11a.InnerText = borderweightleft;
                                //
                                XmlElement nodeN12 = appstyleDoc.CreateElement("borderright");
                                nodeN.AppendChild(nodeN12);
                                nodeN12.InnerText = borderright;
                                XmlElement nodeN12a = appstyleDoc.CreateElement("borderweightright");
                                nodeN.AppendChild(nodeN12a);
                                nodeN12a.InnerText = borderweightright;
                                //
                                XmlElement nodeN13 = appstyleDoc.CreateElement("wraptext");
                                nodeN.AppendChild(nodeN13);
                                nodeN13.InnerText = wraptext;
                                //
                                XmlElement nodeN14 = appstyleDoc.CreateElement("Halignment");
                                nodeN.AppendChild(nodeN14);
                                nodeN14.InnerText = Halignment;
                                //
                                XmlElement nodeN15 = appstyleDoc.CreateElement("Valignment");
                                nodeN.AppendChild(nodeN15);
                                nodeN15.InnerText = Valignment;
                                //
                                XmlElement nodeN16 = appstyleDoc.CreateElement("mergecell");
                                nodeN.AppendChild(nodeN16);
                                nodeN16.InnerText = mergecell;
                                //
                                XmlElement nodeN17 = appstyleDoc.CreateElement("mergecellcount");
                                nodeN.AppendChild(nodeN17);
                                nodeN17.InnerText = mergecellcount;
                                //
                                XmlElement nodeN18 = appstyleDoc.CreateElement("nomberformat");
                                nodeN.AppendChild(nodeN18);
                                nodeN18.InnerText = nomberformat;
                                //
                                XmlElement nodeN19 = appstyleDoc.CreateElement("locked");
                                nodeN.AppendChild(nodeN19);
                                nodeN19.InnerText = locked;
                                //
                                XmlElement nodeN20 = appstyleDoc.CreateElement("formulahidden");
                                nodeN.AppendChild(nodeN20);
                                nodeN20.InnerText = formulahidden;
                                //
                                XmlElement nodeN21 = appstyleDoc.CreateElement("colwidth");
                                nodeN.AppendChild(nodeN21);
                                nodeN21.InnerText = colwidth;
                                //
                                XmlElement nodeN22 = appstyleDoc.CreateElement("rowheight");
                                nodeN.AppendChild(nodeN22);
                                nodeN22.InnerText = rowheight;
                            }

                        }
                        nbstyle.SetAttribute("NB", (nb - 1).ToString());//Total style number
                    }
                    ////////////////////////////////save file dialogue////////////
                    //SaveFileDialog SaveFileDialog2 = new SaveFileDialog();
                    //SaveFileDialog2.InitialDirectory = "D:\\appstyle22.xml";
                    //SaveFileDialog2.Filter = "XML Fichier .xml|*.xml|All files (*.*)|*.*";
                    //SaveFileDialog2.ShowDialog();
                    //string savenom = SaveFileDialog2.FileName.ToString();
                }


                string savenom = textBox2.Text.ToString();


                appstyleDoc.Save(savenom);




                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                int time2 = System.Environment.TickCount;
                int times = time2 - timet;
                string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
                textBox20.AppendText("Travail terminé " + tim + " secondes\r\n" + ". Nom du fichier: " + savenom + System.Environment.NewLine);
                MessageBox.Show("Travail terminé " + tim + " secondes\r\n" + ". Nom du fichier: " + savenom);
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString()+System.Environment.NewLine);
            }

        }


        private void Xmlecriretout_Click(object sender, EventArgs e)
        {
            //OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            //OpenFileDialog1.InitialDirectory = "D:\\ptw\\appstyle.xls";
            //OpenFileDialog1.Filter = "Excel Files .xls|*.xls|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            ////OpenFileDialog1.FilterIndex = 2;
            ////OpenFileDialog1.RestoreDirectory = true;
            //OpenFileDialog1.ShowDialog();
            try
            {
                textBox20.AppendText("==> Start Création du fichier xml des styles" + System.Environment.NewLine);
                ////////////////open excel////////////////////////
                int timet = System.Environment.TickCount;
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet = null;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;

                //string openfilex = OpenFileDialog1.FileName.ToString();//"D:\\ptw\\appstyle.xls"
                string openfilex = textBox7.Text.ToString();
                xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                // xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


                XmlDocument appstyleDoc = new XmlDocument();
                XmlElement styleRoot;
                XmlNode stylexmlNode;

                stylexmlNode = appstyleDoc.CreateNode(XmlNodeType.XmlDeclaration, "", "");
                appstyleDoc.AppendChild(stylexmlNode);
                styleRoot = appstyleDoc.CreateElement("style");
                appstyleDoc.AppendChild(styleRoot);


                ///////////////predifinie palette index////////////////////////////////////////////
                XmlNode racinestyle = appstyleDoc.SelectSingleNode("//style");
                XmlElement xpalette = appstyleDoc.CreateElement("palette");
                racinestyle.AppendChild(xpalette);

                XmlElement nbstyle = appstyleDoc.CreateElement("nbstyle");
                racinestyle.AppendChild(nbstyle);

                for (int nindex = 1; nindex <= 56; nindex++)
                {
                    XmlElement indexx = appstyleDoc.CreateElement("index" + nindex);
                    xpalette.AppendChild(indexx);
                    indexx.InnerText = xlWorkBook.get_Colors(nindex).ToString();
                    ////////////////////////////////////////////attribut RGB faculté
                    int colorindexox = Convert.ToInt32(indexx.InnerText);
                    int colorindexB = colorindexox / 65536;
                    int colorindexG = (colorindexox % 65536) / 256;
                    int colorindexR = (colorindexox % 65536) % 256;
                    indexx.SetAttribute("R", colorindexR.ToString());
                    indexx.SetAttribute("G", colorindexG.ToString());
                    indexx.SetAttribute("B", colorindexB.ToString());
                }
                ///////////////////////////////////////////////////////////////////////////////////

                //routine pour multiple onglet
                //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
                int wkbookcount = xlWorkBook.Worksheets.Count;
                int nb = 1;
                int time1 = System.Environment.TickCount;
                for (int k = 1; k <= wkbookcount - 1; k++)
                {

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(k);

                    Excel.Range range = xlWorkSheet.UsedRange;
                    object[,] values = (object[,])range.Value2;



                    int rCnt = 0;
                    int cCnt = 0;
                    int col = 0;
                    rCnt = range.Rows.Count;
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                        if (Regex.Equals(valuecellabs, "1000"))
                        {
                            col = cCnt;
                            break;
                        }
                    }
                    string colcount = "";


                    for (int row = 1; row <= values.GetUpperBound(0) - 1; row++)
                    {
                        string value = Convert.ToString(values[row, col]);
                        if (value != "")
                        {
                            XmlElement numerostyle = appstyleDoc.CreateElement("nbstyle" + nb);
                            nb++;
                            nbstyle.AppendChild(numerostyle);
                            numerostyle.InnerText = value;
                            ////////////////////////////////////////////////////////////////////
                            XmlElement styleNX = appstyleDoc.CreateElement("style" + value);
                            racinestyle.AppendChild(styleNX);
                            XmlElement colNX = appstyleDoc.CreateElement("col");
                            colNX.InnerText = Convert.ToString(values[row, col + 1]);
                            styleNX.AppendChild(colNX);


                            XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + value);
                            if (xstyle != null)
                            {
                                colcount = (xstyle.SelectSingleNode("col")).InnerText;
                            }

                            //MessageBox.Show(colcount);
                            int colcountx = Convert.ToInt32(colcount);
                            for (int colc = 1; colc <= colcountx; colc++)
                            {
                                XmlElement nodeN = appstyleDoc.CreateElement("style" + value + "." + colc);
                                xstyle.AppendChild(nodeN);

                                Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc + 2] as Excel.Range;
                                string fontname = rangeDelx.Font.Name.ToString();
                                string fontsize = rangeDelx.Font.Size.ToString();
                                string fontcolor = rangeDelx.Font.Color.ToString();
                                int fontcnumber = Convert.ToInt32(fontcolor);
                                int colorB = fontcnumber / 65536;
                                int colorG = (fontcnumber % 65536) / 256;
                                int colorR = (fontcnumber % 65536) % 256;
                                string fontcolorindex = rangeDelx.Font.ColorIndex.ToString();
                                string fontbold = rangeDelx.Font.Bold.ToString();
                                string fontitalic = rangeDelx.Font.Italic.ToString();
                                string fontunderline = rangeDelx.Font.Underline.ToString();
                                string bgcolor = rangeDelx.Interior.Color.ToString();
                                int bgcnumber = Convert.ToInt32(bgcolor);
                                int bgcolorB = bgcnumber / 65536;
                                int bgcolorG = (bgcnumber % 65536) / 256;
                                int bgcolorR = (bgcnumber % 65536) % 256;

                                string bgcolorindex = rangeDelx.Interior.ColorIndex.ToString();
                                //top
                                string bordertop = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle.ToString();
                                string borderweighttop = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight.ToString();
                                //bot
                                string borderbot = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle.ToString();
                                string borderweightbot = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight.ToString();
                                //left
                                string borderleft = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle.ToString();
                                string borderweightleft = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight.ToString();
                                //right
                                string borderright = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle.ToString();
                                string borderweightright = rangeDelx.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight.ToString();

                                string wraptext = rangeDelx.WrapText.ToString();
                                string Halignment = rangeDelx.HorizontalAlignment.ToString();
                                string Valignment = rangeDelx.VerticalAlignment.ToString();
                                string mergecell = rangeDelx.MergeCells.ToString();
                                string mergecellcount = rangeDelx.MergeArea.Cells.Count.ToString();

                                string nomberformat = rangeDelx.NumberFormat.ToString();

                                string locked = rangeDelx.Locked.ToString();
                                string formulahidden = rangeDelx.FormulaHidden.ToString();

                                string colwidth = rangeDelx.ColumnWidth.ToString();
                                string rowheight = rangeDelx.RowHeight.ToString();

                                //
                                XmlElement nodeN1 = appstyleDoc.CreateElement("font");
                                nodeN.AppendChild(nodeN1);
                                nodeN1.InnerText = fontname;
                                //
                                XmlElement nodeN2 = appstyleDoc.CreateElement("fontsize");
                                nodeN.AppendChild(nodeN2);
                                nodeN2.InnerText = fontsize;
                                //
                                XmlElement nodeN3 = appstyleDoc.CreateElement("fontcolorindex");
                                nodeN.AppendChild(nodeN3);
                                nodeN3.InnerText = fontcolorindex;
                                //
                                XmlElement nodeN3a = appstyleDoc.CreateElement("fontcolor");
                                nodeN.AppendChild(nodeN3a);
                                nodeN3a.InnerText = fontcolor;
                                nodeN3a.SetAttribute("R", colorR.ToString());
                                nodeN3a.SetAttribute("G", colorG.ToString());
                                nodeN3a.SetAttribute("B", colorB.ToString());
                                //
                                XmlElement nodeN5 = appstyleDoc.CreateElement("fontbold");
                                nodeN.AppendChild(nodeN5);
                                nodeN5.InnerText = fontbold;
                                //
                                XmlElement nodeN6 = appstyleDoc.CreateElement("fontitalic");
                                nodeN.AppendChild(nodeN6);
                                nodeN6.InnerText = fontitalic;
                                //
                                XmlElement nodeN7 = appstyleDoc.CreateElement("fontunderline");
                                nodeN.AppendChild(nodeN7);
                                nodeN7.InnerText = fontunderline;
                                //
                                XmlElement nodeN8 = appstyleDoc.CreateElement("bgcolor");
                                nodeN.AppendChild(nodeN8);
                                nodeN8.InnerText = bgcolor;
                                nodeN8.SetAttribute("R", bgcolorR.ToString());
                                nodeN8.SetAttribute("G", bgcolorG.ToString());
                                nodeN8.SetAttribute("B", bgcolorB.ToString());
                                //
                                XmlElement nodeN8a = appstyleDoc.CreateElement("bgcolorindex");
                                nodeN.AppendChild(nodeN8a);
                                nodeN8a.InnerText = bgcolorindex;
                                //
                                XmlElement nodeN9 = appstyleDoc.CreateElement("bordertop");
                                nodeN.AppendChild(nodeN9);
                                nodeN9.InnerText = bordertop;
                                XmlElement nodeN9a = appstyleDoc.CreateElement("borderweighttop");
                                nodeN.AppendChild(nodeN9a);
                                nodeN9a.InnerText = borderweighttop;
                                //
                                XmlElement nodeN10 = appstyleDoc.CreateElement("borderbot");
                                nodeN.AppendChild(nodeN10);
                                nodeN10.InnerText = borderbot;
                                XmlElement nodeN10a = appstyleDoc.CreateElement("borderweightbot");
                                nodeN.AppendChild(nodeN10a);
                                nodeN10a.InnerText = borderweightbot;
                                //
                                XmlElement nodeN11 = appstyleDoc.CreateElement("borderleft");
                                nodeN.AppendChild(nodeN11);
                                nodeN11.InnerText = borderleft;
                                XmlElement nodeN11a = appstyleDoc.CreateElement("borderweightleft");
                                nodeN.AppendChild(nodeN11a);
                                nodeN11a.InnerText = borderweightleft;
                                //
                                XmlElement nodeN12 = appstyleDoc.CreateElement("borderright");
                                nodeN.AppendChild(nodeN12);
                                nodeN12.InnerText = borderright;
                                XmlElement nodeN12a = appstyleDoc.CreateElement("borderweightright");
                                nodeN.AppendChild(nodeN12a);
                                nodeN12a.InnerText = borderweightright;
                                //
                                XmlElement nodeN13 = appstyleDoc.CreateElement("wraptext");
                                nodeN.AppendChild(nodeN13);
                                nodeN13.InnerText = wraptext;
                                //
                                XmlElement nodeN14 = appstyleDoc.CreateElement("Halignment");
                                nodeN.AppendChild(nodeN14);
                                nodeN14.InnerText = Halignment;
                                //
                                XmlElement nodeN15 = appstyleDoc.CreateElement("Valignment");
                                nodeN.AppendChild(nodeN15);
                                nodeN15.InnerText = Valignment;
                                //
                                XmlElement nodeN16 = appstyleDoc.CreateElement("mergecell");
                                nodeN.AppendChild(nodeN16);
                                nodeN16.InnerText = mergecell;
                                //
                                XmlElement nodeN17 = appstyleDoc.CreateElement("mergecellcount");
                                nodeN.AppendChild(nodeN17);
                                nodeN17.InnerText = mergecellcount;
                                //
                                XmlElement nodeN18 = appstyleDoc.CreateElement("nomberformat");
                                nodeN.AppendChild(nodeN18);
                                nodeN18.InnerText = nomberformat;
                                //
                                XmlElement nodeN19 = appstyleDoc.CreateElement("locked");
                                nodeN.AppendChild(nodeN19);
                                nodeN19.InnerText = locked;
                                //
                                XmlElement nodeN20 = appstyleDoc.CreateElement("formulahidden");
                                nodeN.AppendChild(nodeN20);
                                nodeN20.InnerText = formulahidden;
                                //
                                XmlElement nodeN21 = appstyleDoc.CreateElement("colwidth");
                                nodeN.AppendChild(nodeN21);
                                nodeN21.InnerText = colwidth;
                                //
                                XmlElement nodeN22 = appstyleDoc.CreateElement("rowheight");
                                nodeN.AppendChild(nodeN22);
                                nodeN22.InnerText = rowheight;
                            }

                        }
                        nbstyle.SetAttribute("NB", (nb - 1).ToString());//Total style number
                    }
                    ////////////////////////////////save file dialogue////////////
                    //SaveFileDialog SaveFileDialog2 = new SaveFileDialog();
                    //SaveFileDialog2.InitialDirectory = "D:\\appstyle22.xml";
                    //SaveFileDialog2.Filter = "XML Fichier .xml|*.xml|All files (*.*)|*.*";
                    //SaveFileDialog2.ShowDialog();
                    //string savenom = SaveFileDialog2.FileName.ToString();
                }


                string savenom = textBox2.Text.ToString();


                appstyleDoc.Save(savenom);




                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                int time2 = System.Environment.TickCount;
                int times = time2 - timet;
                string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
                textBox20.AppendText("Travail terminé " + tim + " secondes\r\n" + ". Nom du fichier: " + savenom + System.Environment.NewLine);
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
        }



        //
        ////Xml lire Xmllire_Click
        //
        //private void Xmllire_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
        //    OpenFileDialog1.FileName = fileAstyler;
        //    OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
        //    OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
        //    //OpenFileDialog1.FilterIndex = 2;
        //    OpenFileDialog1.RestoreDirectory = true;
        //    if (OpenFileDialog1.FileName == "")
        //    {
        //        OpenFileDialog1.FileName = textBox14.Text.ToString();
        //        //OpenFileDialog1.ShowDialog();
        //    }

        //    ////////////////open excel////////////////////////
        //    Excel.Application xlApp2;
        //    Excel.Workbook xlWorkBook;
        //    object misValue = System.Reflection.Missing.Value;
        //    xlApp2 = new Excel.ApplicationClass();
        //    xlApp2.Visible = true;
        //    xlApp2.DisplayAlerts = false;

        //    Thread.Sleep(3000);
        //    string openfilex = OpenFileDialog1.FileName.ToString();



        //    try
        //    {


        //        Excel.Workbook xlworkbookStyle = xlApp2.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");
        //        Excel.Workbook xlworkbookNP = xlApp2.Workbooks.Open(openfilex);

        //        Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookStyle.Worksheets.get_Item("Histo et Histo-s");
        //        Excel._Worksheet xlworksheet = (Excel.Worksheet)xlworkbookNP.Worksheets.get_Item("Historique");

        //        //Excel.Range range= xlworksheet.get_Range("A15","M1856");
        //        Excel.Range rangeAll = xlworksheet.UsedRange;
        //        rangeAll.ClearFormats();
        //        rangeAll.Interior.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 0, 0));

        //        //xlworkbookNP.SaveAs("D:\\ptw\\changeStyle\\123.xlsx");
        //        object[,] values1 = (object[,])rangeAll.Value2;
        //        Excel.Range rangeUsedStyle = xlworksheetStyle.UsedRange;
        //        int col17000style = 0;
        //        object[,] valueStyle = (object[,])rangeUsedStyle.Value2;
        //        for (int i = 1; i <= rangeUsedStyle.Columns.Count; i++)
        //        {
        //            if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() != null)
        //            {
        //                if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() == "17000")
        //                {
        //                    col17000style = i;
        //                    break;
        //                }
        //            }
        //        }


        //        for (int i = 14; i <= 79; i++)
        //        {
        //            Excel.Range chercheStyle = xlworksheetStyle.get_Range("A" + i, "A" + i);
        //            Excel.Range changeFontRange = xlworksheetStyle.Cells[i, col17000style] as Excel.Range;
        //            if (changeFontRange.Value2 != null)
        //            {
        //                int fontToChange = int.Parse(changeFontRange.Value2.ToString());

        //                Excel.Range rangeToChangeFont = xlworksheetStyle.get_Range("C" + i, "O" + i);
        //               // rangeToChangeFont.EntireRow.Font.Size = fontToChange;




        //            }
        //            string cherche = "";
        //            if (chercheStyle.Value2 != null)
        //            {
        //                cherche = chercheStyle.Value2.ToString();
        //            }
        //            else
        //            {
        //                continue;
        //            }
        //            Excel.Range rangeStyle = xlworksheetStyle.get_Range("C" + i, "O" + i);
        //            rangeStyle.Copy();

        //            int row8180002000 = rangeAll.Rows.Count;

        //            rangeAll.get_Range(xlworksheet.Cells[1,14], xlworksheet.Cells[1, xlworksheet.UsedRange.Columns.Count]).EntireColumn.Hidden = true;
                    
        //            for (int t = 1; t <= row8180002000; t++)
        //            {

        //                Excel.Range rangeColN = xlworksheet.get_Range("N" + t, "N" + t);
        //                if (rangeColN.Value2 != null)
        //                {
        //                    string x = rangeColN.Value2.ToString();
        //                    if (cherche == "12000")
        //                    {
        //                        int sd = 0;
        //                        string xss = rangeColN.Value2.ToString();
        //                    }
        //                    if (rangeColN.Value2.ToString() == cherche)
        //                    {

        //                        Excel.Range rangePasteStyle = xlworksheet.get_Range("A" + t, "M" + t);
        //                        rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationSubtract, false, false);
        //                        //rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        //                        rangePasteStyle.EntireRow.Font.Size = rangeStyle.Font.Size;
        //                        if (rangeColN.Value2.ToString() == "4000-100" || rangeColN.Value2.ToString() == "12000-1000" || rangeColN.Value2.ToString() == "10000-100" || rangeColN.Value2.ToString() == "10000-200" || rangeColN.Value2.ToString() == "10000")
        //                        {
        //                            Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
        //                            rangeAutoFit.EntireRow.RowHeight = 7; ;
        //                        }//------------------------hide some lines--------------------------------------------
        //                        else if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
        //                        {
        //                            Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
        //                            rangeToHide.Hidden = true;
        //                        }
        //                        //------------------------------------------------------------------------------------


        //                            //------------------------auto fit the height of the big title rows------------------------
        //                        else if (rangeColN.Value2.ToString() == "6000" || rangeColN.Value2.ToString() == "5000")
        //                        {
        //                            Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
        //                            rangeAutoFit.EntireRow.AutoFit();
        //                        }
        //                    }
        //                    if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
        //                    {
        //                        Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
        //                        rangeToHide.Hidden = true;
        //                    }
        //                    if (rangeColN.Value2.ToString() == "")
        //                    {
        //                        rangeColN.EntireRow.Hidden = true;
        //                    }
        //                }
        //                else
        //                {
        //                    rangeColN.EntireRow.Hidden = true;
        //                }
        //            }
        //        }
        //        if (divitylerfinal != null) xlworkbookNP.SaveCopyAs(divitylerfinal);
        //        xlApp2.Quit();
               
        //    }
        //    catch (Exception ex)
        //    {
        //        xlApp2.Quit();
        //        textBox20.AppendText(ex.ToString());
        //    }



        //    //xlWorkBook = xlApp2.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //    ////xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
        //    ////Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
        //    //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];
        //    //Excel.Range range = xlWorkSheet.UsedRange;
        //    //object[,] values = (object[,])range.Value2;

        //    ////////////////////////////////open le fichier style XML//////////////////////
        //    ////OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
        //    ////OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
        //    ////OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
        //    ////OpenFileDialog2.ShowDialog();
        //    ////stylexml = OpenFileDialog2.FileName.ToString();

        //    //if (textBox10.Text != null)
        //    //    stylexml = textBox10.Text;
        //    //else
        //    //    MessageBox.Show("Veuillez choiser le fichier style en format XML");


        //    //XmlDocument appstyleDoc = new XmlDocument();
        //    //appstyleDoc.Load(stylexml);
        //    ////appstyleDoc.Load("D:\\appstyle22.xml");

        //    ///////////////////////////////////////set palette couleur///////////////////////////
        //    ////xlWorkBook.ResetColors();
        //    //XmlElement indexxmlelement = appstyleDoc.DocumentElement;
        //    //XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
        //    //XmlNode indexstylenode = indexstylenodelist.Item(0);

        //    ////!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!pour test2010
        //    ////for (int nindex = 1; nindex <= 56; nindex++)
        //    ////{
        //    ////    string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
        //    ////    int valeur2index = Convert.ToInt32(valeurindex);
        //    ////    xlWorkBook.set_Colors(nindex, valeur2index);
        //    ////}


        //    //range.EntireRow.Font.Size = 8;
        //    ////range.Rows.AutoFit();
        //    ////Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //    ////rangemasquer.EntireRow.Hidden = true;
        //    //////////////////////////////////////////////////////////////////////////////////////

        //    //int rCnt = 0;
        //    //int cCnt = 0;
        //    ////int col = 0;
        //    //int col15000 = 0;
        //    //int colannuel9000 = 0;
        //    //rCnt = range.Rows.Count;

        //    //int col83000 = 0;
        //    //int col8000 = 0;

        //    //CodeFinder cf;
        //    //cf = new CodeFinder(xlWorkBook, xlWorkSheet);
        //    //col15000 = cf.FindCodedColumn("6000", range);
        //    ////colannuel9000 = cf.FindCodedColumn("9000-1000", range);
        //    //col8000 = cf.FindCodedColumn("8000", range);
        //    //col83000 = cf.FindCodedColumn("83000", range);


        //    ////for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
        //    ////{
        //    ////    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
        //    ////    if (Regex.Equals(valuecellabs, "9000-1000"))
        //    ////    {
        //    ////        colannuel9000 = cCnt;
        //    ////    }
        //    ////    if (Regex.Equals(valuecellabs, "8000"))
        //    ////    {
        //    ////        col8000 = cCnt;
        //    ////    }
        //    ////    if (Regex.Equals(valuecellabs, "83000"))
        //    ////    {
        //    ////        col83000 = cCnt;
        //    ////        break;
        //    ////    }
        //    ////}


            
        //    /////////////////////////////////////construit tableaux style///////////MAX 100///////////////
        //    //XmlElement nbstyle = appstyleDoc.DocumentElement;
        //    //XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
        //    //XmlNode nbstylenode = nbstylelist.Item(0);

        //    //string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
        //    //int nbtotalint = 120;
        //    //nbtotalint = Convert.ToInt32(nbtotal);
        //    //string[] tablestyle = new string[nbtotalint+1];
        //    //for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //    //{
        //    //    tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
        //    //}
        //    ///////////////////////////////////////////////////////////////////////////////////////////////


        //    //int row = 1;
        //    //string colcount = "";
        //    //int time1 = System.Environment.TickCount;
        //    //int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
        //    //for (row = 1; row <= rowCountx-1; row++)
        //    //{
        //    //    string value = Convert.ToString(values[row, col15000]);
        //    //    for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //    //    {
        //    //        if (Regex.Equals(value, tablestyle[nbs]))
        //    //        {
        //    //            XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
        //    //            if (xstyle != null)
        //    //            {
        //    //                colcount = (xstyle.SelectSingleNode("col")).InnerText;
        //    //            }
        //    //            int colcountx = Convert.ToInt32(colcount);
        //    //            for (int colc = 1; colc <= colcountx; colc++)
        //    //            {
        //    //                XmlElement xmlelement = appstyleDoc.DocumentElement;
        //    //                XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
        //    //                XmlNode stylenode = stylenodelist.Item(0);
        //    //                string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
        //    //                string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
        //    //                //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
        //    //                //int colorBx = Convert.ToInt32(colorB);
        //    //                //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
        //    //                string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
        //    //                int fcolor = Convert.ToInt32(fontcolor);
        //    //                //string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

        //    //                string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
        //    //                string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
        //    //                string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

        //    //                string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
        //    //                //int bcolor = Convert.ToInt32(bgcolor);
        //    //                string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
        //    //                string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
        //    //                string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
        //    //                string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
        //    //                string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
        //    //                string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
        //    //                string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
        //    //                string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
        //    //                string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

        //    //                string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
        //    //                string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
        //    //                string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
        //    //                string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
        //    //                string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //    //                int intmergecellcount = Convert.ToInt32(mergecellcount);

        //    //                string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
        //    //                string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
        //    //                string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
        //    //                string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
        //    //                string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
        //    //                ///////////////////////////////////merge process///////////////////////////////////////////
        //    //                if (mergecell == "True")
        //    //                {
        //    //                    if (intmergecellcount > 1)
        //    //                    {
        //    //                        Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
        //    //                        rangemerge.Merge(false);
        //    //                        //rangemerge.HorizontalAlignment = 1;

        //    //                        for (int countarea = 1; countarea < intmergecellcount; countarea++)
        //    //                        {
        //    //                            XmlElement mergexmlelement = appstyleDoc.DocumentElement;
        //    //                            int mergecolindex = colc + countarea;
        //    //                            XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
        //    //                            XmlNode mergestylenode = mergestylenodelist.Item(0);
        //    //                            mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
        //    //                            appstyleDoc.Save(stylexml);
        //    //                        }
        //    //                    }
        //    //                }
        //    //                /////////////////////////////exception traitement/////////////////////////////////
        //    //                //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
        //    //                //xlWorkSheet.Cells.ColumnWidth = 20;
        //    //                //////////////////////////////////////////////////////////////////////////////////

        //    //                /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
        //    //                Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
        //    //                rangeDelx.Font.Name = fontname;
        //    //                rangeDelx.Font.Size = Convert.ToInt32(fontsize);
        //    //               // rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
        //    //                rangeDelx.Font.Color = fcolor;


        //    //                rangeDelx.Font.Bold = (fontbold=="True");
        //    //                rangeDelx.Font.Italic = (fontitalic == "True");
        //    //                rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
        //    //                //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);
        //    //                rangeDelx.Interior.Color = bgcolor;
        //    //               // rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
        //    //                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);
                            
        //    //                rangeDelx.WrapText = (wraptext == "True");
        //    //                rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
        //    //                rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

        //    //                /////////////////////////////////////////////////////////////////////////////////////////
        //    //                mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //    //                //ne peut pas modifier les cellules fusionner
        //    //                if (mergecellcount == "False")
        //    //                {
        //    //                    rangeDelx.NumberFormat = nomberformat;
        //    //                    rangeDelx.Locked = (locked == "True");
        //    //                    rangeDelx.Locked = (formulahidden == "True");
        //    //                }
        //    //                ///////////////////////////////////////////////////////////////////////////////////////////
        //    //                rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
        //    //                rangeDelx.RowHeight = Convert.ToDouble(rowheight);
        //    //            }
        //    //        }
        //    //    }
        //    //}
        //    //xlApp2.ActiveWindow.DisplayGridlines = false;
        //    ////range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        //    ////range.Rows.AutoFit();
        //    ////Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //    ////rangemasquer2.EntireRow.Hidden = true;

        //    ////pour consigne de masquage
        //    //Excel.Range rangeremplace = xlWorkSheet.UsedRange;
        //    //object[,] values8000 = (object[,])rangeremplace.Value2;
        //    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
        //    //{
        //    //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
        //    //    if (Regex.Equals(valuedel, "-1"))
        //    //    {
        //    //        Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
        //    //        rangeDely.EntireRow.Hidden = true;
        //    //    }
        //    //}
        //    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
        //    //{
        //    //    string valuedel = Convert.ToString(values8000[rowhide, col8000]);
        //    //    if (Regex.Equals(valuedel, "-5"))
        //    //    {
        //    //        Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col8000] as Excel.Range;
        //    //        rangeDely.EntireRow.Hidden = true;
        //    //    }
        //    //}

        //    //Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("N1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
        //    //Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
        //    //Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;
        //    ////consigne supression
        //    ////rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
        //    ////rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //    ////rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //    //rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
        //    //rangeDelete.EntireColumn.Hidden = true;

        //    //Excel.Worksheet hisrefer;
        //    //Excel.Range referrange;
        //    //if (openfilex.Contains("-s"))
        //    //{
        //    //    hisrefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");
        //    //    referrange = hisrefer.UsedRange;
        //    //    referrange.Copy();
        //    //}
        //    //else
        //    //{
        //    //    hisrefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
        //    //    referrange = hisrefer.UsedRange;
        //    //    referrange.Copy();
        //    //}
        //    //string smallfile = openfilex.Substring(12).Split('.')[0];
        //    //int referrow = referrange.Rows.Count;
        //    //int refercol = referrange.Columns.Count;
        //    //object[,] refervalue = (object[,])referrange.Value2;

        //    //for (int i = 1; i < refercol; i++)
        //    //{
        //    //    if (refervalue[1, i] != null)
        //    //    {
        //    //        if (refervalue[1, i].ToString() == smallfile)
        //    //        {

        //    //            int a = 0;
        //    //            for (int j = 1; j < referrow - a; j++)
        //    //            {
        //    //                if (refervalue[j, i] == null)
        //    //                {

        //    //                    Excel.Range deletereferrow = hisrefer.Cells[j, i] as Excel.Range;
        //    //                    //deletereferrow.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //    //                    deletereferrow.Clear();
        //    //                    //j--;
        //    //                    //a++;
        //    //                    referrange = hisrefer.UsedRange;
        //    //                    refervalue = (object[,])referrange.Value2;
        //    //                }
        //    //            }
        //    //        }
        //    //    }
        //    //}
            
        //    //int time2 = System.Environment.TickCount;
        //    //int times = time2 - time1;
        //    //string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
        //    ////MessageBox.Show("jobs done " + tim + " seconds used");
        //    ////xlWorkBook.Save();
        //    //if (pathstylerfinal != null) System.IO.Directory.CreateDirectory(pathstylerfinal);
        //    //if (divitylerfinal != null) xlWorkBook.SaveCopyAs(divitylerfinal);
        //    //if (divitylerfinal != null) xlWorkBook.Close(false, misValue, misValue);

        //    //if (divitylerfinal == null) xlWorkBook.Close(true, misValue, misValue);
        //    //xlApp2.Quit();

        //    //releaseObject(xlWorkSheet);
        //    //releaseObject(xlWorkBook);
        //    //releaseObject(xlApp2);
        //}


        //
        ////Styler annuel.ptw
        //
        //
        ////Xml lire Xmllire_Click
        //
        private void Xmllire_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.FileName = fileAstyler;
            OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
            OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            //OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;
            if (OpenFileDialog1.FileName == "")
            {
                OpenFileDialog1.FileName = textBox14.Text.ToString();
                //OpenFileDialog1.ShowDialog();
            }

            ////////////////open excel////////////////////////
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp2 = new Excel.ApplicationClass();
            xlApp2.Visible = true;
            xlApp2.DisplayAlerts = false;

            Thread.Sleep(3000);
            string openfilex = OpenFileDialog1.FileName.ToString();
            xlWorkBook = xlApp2.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////
            //OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
            //OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
            //OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
            //OpenFileDialog2.ShowDialog();
            //stylexml = OpenFileDialog2.FileName.ToString();

            if (textBox10.Text != null)
                stylexml = textBox10.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");


            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);
            //appstyleDoc.Load("D:\\appstyle22.xml");

            /////////////////////////////////////set palette couleur///////////////////////////
            //xlWorkBook.ResetColors();
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);

            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!pour test2010
            //for (int nindex = 1; nindex <= 56; nindex++)
            //{
            //    string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
            //    int valeur2index = Convert.ToInt32(valeurindex);
            //    xlWorkBook.set_Colors(nindex, valeur2index);
            //}


            range.EntireRow.Font.Size = 8;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer.EntireRow.Hidden = true;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;
            int col15000 = 0;
            int colannuel9000 = 0;
            rCnt = range.Rows.Count;

            int col83000 = 0;
            int col8000 = 0;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col15000 = cf.FindCodedColumn("6000", range);
            //colannuel9000 = cf.FindCodedColumn("9000-1000", range);
            col8000 = cf.FindCodedColumn("8000", range);
            col83000 = cf.FindCodedColumn("83000", range);


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "9000-1000"))
            //    {
            //        colannuel9000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "8000"))
            //    {
            //        col8000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "83000"))
            //    {
            //        col83000 = cCnt;
            //        break;
            //    }
            //}



            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint + 1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////


            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;
            int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
            for (row = 1; row <= rowCountx - 1; row++)
            {
                string value = Convert.ToString(values[row, col15000]);
                for (int nbs = 1; nbs <= nbtotalint; nbs++)
                {
                    if (Regex.Equals(value, tablestyle[nbs]))
                    {
                        XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                        if (xstyle != null)
                        {
                            colcount = (xstyle.SelectSingleNode("col")).InnerText;
                        }
                        int colcountx = Convert.ToInt32(colcount);
                        for (int colc = 1; colc <= colcountx; colc++)
                        {
                            XmlElement xmlelement = appstyleDoc.DocumentElement;
                            XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                            XmlNode stylenode = stylenodelist.Item(0);
                            string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                            string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                            //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                            //int colorBx = Convert.ToInt32(colorB);
                            //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                            string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                            int fcolor = Convert.ToInt32(fontcolor);
                            //string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                            string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                            string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                            string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                            string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
                            //int bcolor = Convert.ToInt32(bgcolor);
                            string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                            string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                            string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                            string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                            string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                            string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                            string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                            string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                            string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                            string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                            string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                            string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                            string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                            string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            int intmergecellcount = Convert.ToInt32(mergecellcount);

                            string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                            string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                            string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                            string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                            string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
                            ///////////////////////////////////merge process///////////////////////////////////////////
                            if (mergecell == "True")
                            {
                                if (intmergecellcount > 1)
                                {
                                    Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
                                    rangemerge.Merge(false);
                                    //rangemerge.HorizontalAlignment = 1;

                                    for (int countarea = 1; countarea < intmergecellcount; countarea++)
                                    {
                                        XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                        int mergecolindex = colc + countarea;
                                        XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                        XmlNode mergestylenode = mergestylenodelist.Item(0);
                                        mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                        appstyleDoc.Save(stylexml);
                                    }
                                }
                            }
                            /////////////////////////////exception traitement/////////////////////////////////
                            //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                            //xlWorkSheet.Cells.ColumnWidth = 20;
                            //////////////////////////////////////////////////////////////////////////////////

                            /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
                            Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                            rangeDelx.Font.Name = fontname;
                            rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                            // rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                            rangeDelx.Font.Color = fcolor;


                            rangeDelx.Font.Bold = (fontbold == "True");
                            rangeDelx.Font.Italic = (fontitalic == "True");
                            rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                            //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);
                            rangeDelx.Interior.Color = bgcolor;
                            // rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                            rangeDelx.WrapText = (wraptext == "True");
                            rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                            rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

                            /////////////////////////////////////////////////////////////////////////////////////////
                            mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            //ne peut pas modifier les cellules fusionner
                            if (mergecellcount == "1")
                            {
                               
                                    rangeDelx.NumberFormat = nomberformat;
                                    try
                                    {
                                        rangeDelx.Locked = (locked == "True");
                                        rangeDelx.Locked = (formulahidden == "True");
                                    }
                                    catch
                                    {
                                    }
                            }
                            ///////////////////////////////////////////////////////////////////////////////////////////
                            rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                            rangeDelx.RowHeight = Convert.ToDouble(rowheight);
                        }
                    }
                }
            }
            xlApp2.ActiveWindow.DisplayGridlines = false;
            //range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer2.EntireRow.Hidden = true;

            //pour consigne de masquage
            Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            object[,] values8000 = (object[,])rangeremplace.Value2;
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                if (Regex.Equals(valuedel, "-1"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                if (Regex.Equals(valuedel, "-5"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col8000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }

            Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("N1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
            Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
            Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;
            //consigne supression
            //rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            //rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
            rangeDelete.EntireColumn.Hidden = true;

            Excel.Worksheet hisrefer;
            Excel.Range referrange;
            if (openfilex.Contains("-s"))
            {
                hisrefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");
                referrange = hisrefer.UsedRange;
                referrange.Copy();
            }
            else
            {
                hisrefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
                referrange = hisrefer.UsedRange;
                referrange.Copy();
            }
            string smallfile = openfilex.Substring(12).Split('.')[0];
            int referrow = referrange.Rows.Count;
            int refercol = referrange.Columns.Count;
            object[,] refervalue = (object[,])referrange.Value2;

            for (int i = 1; i < refercol; i++)
            {
                if (refervalue[1, i] != null)
                {
                    if (refervalue[1, i].ToString() == smallfile)
                    {

                        int a = 0;
                        for (int j = 1; j < referrow - a; j++)
                        {
                            if (refervalue[j, i] == null)
                            {

                                Excel.Range deletereferrow = hisrefer.Cells[j, i] as Excel.Range;
                                //deletereferrow.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                deletereferrow.Clear();
                                //j--;
                                //a++;
                                referrange = hisrefer.UsedRange;
                                refervalue = (object[,])referrange.Value2;
                            }
                        }
                    }
                }
            }

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            //xlWorkBook.Save();
            if (pathstylerfinal != null) System.IO.Directory.CreateDirectory(pathstylerfinal);
            if (divitylerfinal != null) xlWorkBook.SaveCopyAs(divitylerfinal);
            if (divitylerfinal != null) xlWorkBook.Close(false, misValue, misValue);

            if (divitylerfinal == null) xlWorkBook.Close(true, misValue, misValue);
            xlApp2.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp2);
        }
        //private void XmllireAnnuel_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
        //    OpenFileDialog1.FileName = fileAstyler;
        //    OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
        //    OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
        //    OpenFileDialog1.FilterIndex = 2;
        //    OpenFileDialog1.RestoreDirectory = true;
        //    if (OpenFileDialog1.FileName == "")
        //    {
        //        OpenFileDialog1.ShowDialog();
        //    }
        //    Thread.Sleep(3000);
        //    //////////////open excel////////////////////////
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    object misValue = System.Reflection.Missing.Value;
        //    xlApp = new Excel.ApplicationClass();
        //    xlApp.Visible = true;
        //    xlApp.DisplayAlerts = false;

            
        //    string openfilex = OpenFileDialog1.FileName.ToString();
        //    xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //    xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
        //    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
        //    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
        //    Excel.Range range = xlWorkSheet.UsedRange;
        //    object[,] values = (object[,])range.Value2;

        //    ////////////////////////////open le fichier style XML//////////////////////
        //    OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
        //    OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
        //    OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
        //    OpenFileDialog2.ShowDialog();
        //    stylexml = OpenFileDialog2.FileName.ToString();

        //    if (textBox2.Text != null)
        //        stylexml = textBox2.Text;
        //    else
        //        MessageBox.Show("Veuillez choiser le fichier style en format XML");


        //    XmlDocument appstyleDoc = new XmlDocument();
        //    appstyleDoc.Load(stylexml);
        //    appstyleDoc.Load("D:\\appstyle22.xml");

        //    ///////////////////////////////////set palette couleur///////////////////////////
        //    xlWorkBook.ResetColors();
        //    XmlElement indexxmlelement = appstyleDoc.DocumentElement;
        //    XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
        //    XmlNode indexstylenode = indexstylenodelist.Item(0);

        //    for (int nindex = 1; nindex <= 56; nindex++)
        //    {
        //        string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
        //        int valeur2index = Convert.ToInt32(valeurindex);
        //        xlWorkBook.set_Colors(nindex, valeur2index);
        //    }
        //    range.EntireRow.Font.Size = 8;
        //    range.Rows.AutoFit();
        //    Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //    rangemasquer.EntireRow.Hidden = true;
        //    //////////////////////////////////////////////////////////////////////////////////

        //    int rCnt = 0;
        //    int cCnt = 0;
        //    int col = 0;

        //    int colannuel9000 = 0;
        //    rCnt = range.Rows.Count;

        //    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
        //    {
        //        string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
        //        if (Regex.Equals(valuecellabs, "9000-1000"))
        //        {
        //            colannuel9000 = cCnt;
        //            break;
        //        }
        //    }



        //    /////////////////////////////////construit tableaux style///////////MAX 100///////////////
        //    XmlElement nbstyle = appstyleDoc.DocumentElement;
        //    XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
        //    XmlNode nbstylenode = nbstylelist.Item(0);

        //    string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
        //    int nbtotalint = 120;
        //    nbtotalint = Convert.ToInt32(nbtotal);
        //    string[] tablestyle = new string[nbtotalint+1];
        //    for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //    {
        //        tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
        //    }
        //    ///////////////////////////////////////////////////////////////////////////////////////////


        //    int row = 1;
        //    string colcount = "";
        //    int time1 = System.Environment.TickCount;

        //    for (row = 1; row <= values.GetUpperBound(0); row++)
        //    {
        //        string value = Convert.ToString(values[row, colannuel9000]);
        //        for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //        {
        //            if (Regex.Equals(value, tablestyle[nbs]))
        //            {
        //                XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
        //                if (xstyle != null)
        //                {
        //                    colcount = (xstyle.SelectSingleNode("col")).InnerText;
        //                }
        //                int colcountx = Convert.ToInt32(colcount);
        //                for (int colc = 1; colc <= colcountx; colc++)
        //                {
        //                    XmlElement xmlelement = appstyleDoc.DocumentElement;
        //                    XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
        //                    XmlNode stylenode = stylenodelist.Item(0);
        //                    string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
        //                    string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
        //                    string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
        //                    int colorBx = Convert.ToInt32(colorB);
        //                    int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
        //                    string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
        //                    string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

        //                    string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
        //                    string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
        //                    string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

        //                    string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
        //                    string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
        //                    string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
        //                    string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
        //                    string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
        //                    string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
        //                    string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
        //                    string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
        //                    string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

        //                    string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
        //                    string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
        //                    string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
        //                    string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
        //                    string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //                    int intmergecellcount = Convert.ToInt32(mergecellcount);

        //                    string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
        //                    string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
        //                    string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
        //                    string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
        //                    string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
        //                    /////////////////////////////////merge process///////////////////////////////////////////
        //                    if (mergecell == "True")
        //                    {
        //                        if (intmergecellcount > 1)
        //                        {
        //                            Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
        //                            rangemerge.Merge(false);
        //                            rangemerge.HorizontalAlignment = 1;

        //                            for (int countarea = 1; countarea < intmergecellcount; countarea++)
        //                            {
        //                                XmlElement mergexmlelement = appstyleDoc.DocumentElement;
        //                                int mergecolindex = colc + countarea;
        //                                XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
        //                                XmlNode mergestylenode = mergestylenodelist.Item(0);
        //                                mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
        //                                appstyleDoc.Save(stylexml);
        //                            }
        //                        }
        //                    }
        //                    ///////////////////////////exception traitement/////////////////////////////////
        //                    Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
        //                    xlWorkSheet.Cells.ColumnWidth = 20;
        //                    ////////////////////////////////////////////////////////////////////////////////

        //                    ///////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
        //                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
        //                    rangeDelx.Font.Name = fontname;
        //                    rangeDelx.Font.Size = Convert.ToInt32(fontsize);
        //                    rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
        //                    rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
        //                    rangeDelx.Value2 = fontcolorindex;

        //                    rangeDelx.Font.Bold = (fontbold == "True");
        //                    rangeDelx.Font.Italic = (fontitalic == "True");
        //                    rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
        //                    rangeDelx.Value2 += "bgcolor" + bgcolorindex;
        //                    rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
        //                    rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

        //                    rangeDelx.WrapText = (wraptext == "True");
        //                    rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
        //                    rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

        //                    ///////////////////////////////////////////////////////////////////////////////////////
        //                    mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //                    ne peut pas modifier les cellules fusionner
        //                    if (mergecellcount == "False")
        //                    {
        //                        rangeDelx.NumberFormat = nomberformat;
        //                        rangeDelx.Locked = (locked == "True");
        //                        rangeDelx.Locked = (formulahidden == "True");
        //                    }
        //                    /////////////////////////////////////////////////////////////////////////////////////////
        //                    rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
        //                    rangeDelx.RowHeight = Convert.ToDouble(rowheight);
        //                }
        //            }
        //        }
        //    }
        //    xlApp.ActiveWindow.DisplayGridlines = false;
        //    range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        //    range.Rows.AutoFit();
        //    Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //    rangemasquer2.EntireRow.Hidden = true;


        //    Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("Y1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
        //    Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
        //    Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;
        //    rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
        //    rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //    rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);


        //    int time2 = System.Environment.TickCount;
        //    int times = time2 - time1;
        //    string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
        //    MessageBox.Show("jobs done " + tim + " seconds used");
        //    xlWorkBook.Save();
        //    System.IO.Directory.CreateDirectory(pathstylerfinal);
        //    xlWorkBook.SaveCopyAs(divitylerfinal);
        //    xlWorkBook.Close(false, misValue, misValue);
        //    xlApp.Quit();

        //    releaseObject(xlWorkSheet);
        //    releaseObject(xlWorkBook);
        //    releaseObject(xlApp);
        //}

        //
        ////Styler annuel.ptw
        //
        private void XmllireAnnuel_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.FileName = fileAstyler;
            OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
            OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            //OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;
            if (OpenFileDialog1.FileName == "")
            {
                OpenFileDialog1.ShowDialog();
            }
            Thread.Sleep(3000);
            ////////////////open excel////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;


            string openfilex = OpenFileDialog1.FileName.ToString();
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            // xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////
            //OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
            //OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
            //OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
            //OpenFileDialog2.ShowDialog();
            //stylexml = OpenFileDialog2.FileName.ToString();

            if (textBox2.Text != null)
                stylexml = textBox2.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");


            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);
            //appstyleDoc.Load("D:\\appstyle22.xml");

            /////////////////////////////////////set palette couleur///////////////////////////
            xlWorkBook.ResetColors();
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);

            for (int nindex = 1; nindex <= 56; nindex++)
            {
                string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
                int valeur2index = Convert.ToInt32(valeurindex);
                xlWorkBook.set_Colors(nindex, valeur2index);
            }
            range.EntireRow.Font.Size = 8;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer.EntireRow.Hidden = true;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;

            int colannuel9000 = 0;
            rCnt = range.Rows.Count;

            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "9000-1000"))
                {
                    colannuel9000 = cCnt;
                    break;
                }
            }



            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint + 1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////


            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;

            for (row = 1; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, colannuel9000]);
                for (int nbs = 1; nbs <= nbtotalint; nbs++)
                {
                    if (Regex.Equals(value, tablestyle[nbs]))
                    {
                        XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                        if (xstyle != null)
                        {
                            colcount = (xstyle.SelectSingleNode("col")).InnerText;
                        }
                        int colcountx = Convert.ToInt32(colcount);
                        for (int colc = 1; colc <= colcountx; colc++)
                        {
                            XmlElement xmlelement = appstyleDoc.DocumentElement;
                            XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                            XmlNode stylenode = stylenodelist.Item(0);
                            string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                            string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                            //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                            //int colorBx = Convert.ToInt32(colorB);
                            //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                            //string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                            string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                            string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                            string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                            string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                            string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                            string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                            string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                            string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                            string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                            string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                            string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                            string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                            string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                            string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                            string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                            string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                            string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                            string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            int intmergecellcount = Convert.ToInt32(mergecellcount);

                            string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                            string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                            string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                            string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                            string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
                            ///////////////////////////////////merge process///////////////////////////////////////////
                            if (mergecell == "True")
                            {
                                if (intmergecellcount > 1)
                                {
                                    Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
                                    rangemerge.Merge(false);
                                    //rangemerge.HorizontalAlignment = 1;

                                    for (int countarea = 1; countarea < intmergecellcount; countarea++)
                                    {
                                        XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                        int mergecolindex = colc + countarea;
                                        XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                        XmlNode mergestylenode = mergestylenodelist.Item(0);
                                        mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                        appstyleDoc.Save(stylexml);
                                    }
                                }
                            }
                            /////////////////////////////exception traitement/////////////////////////////////
                            //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                            //xlWorkSheet.Cells.ColumnWidth = 20;
                            //////////////////////////////////////////////////////////////////////////////////

                            /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
                            Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                            rangeDelx.Font.Name = fontname;
                            rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                            //rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
                            rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                            //rangeDelx.Value2 = fontcolorindex;

                            rangeDelx.Font.Bold = (fontbold == "True");
                            rangeDelx.Font.Italic = (fontitalic == "True");
                            rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                            //rangeDelx.Value2 += "bgcolor" + bgcolorindex;
                            rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                            rangeDelx.WrapText = (wraptext == "True");
                            rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                            rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

                            /////////////////////////////////////////////////////////////////////////////////////////
                            mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            //ne peut pas modifier les cellules fusionner
                            if (mergecellcount == "1")
                            {
                                rangeDelx.NumberFormat = nomberformat;
                                try
                                {
                                    rangeDelx.Locked = (locked == "True");
                                    rangeDelx.Locked = (formulahidden == "True");
                                }
                                catch
                                {
                                }
                            }
                            ///////////////////////////////////////////////////////////////////////////////////////////
                            rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                            rangeDelx.RowHeight = Convert.ToDouble(rowheight);
                        }
                    }
                }
            }
            xlApp.ActiveWindow.DisplayGridlines = false;
            //range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer2.EntireRow.Hidden = true;


            Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("Y1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
            Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
            //Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;
            rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            //xlWorkBook.Save();
            System.IO.Directory.CreateDirectory(pathstylerfinal);
            xlWorkBook.SaveCopyAs(divitylerfinal);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        #endregion

        #region Fonctions pour les buttons

        private void Parcourir_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox1.Text = OpenFileDialogx.FileName;
        }
        private void button8_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox12.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }
        //Repertoire source pour le fichier à subdiviser
        private void button27_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox2.Text = OpenFileDialogx.FileName;
        }
        //Repertoire Destitation pour les fichier subdiviser
        private void button28_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox3.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox2.Text = OpenFileDialogx.FileName;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox7.Text = OpenFileDialogx.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox8.Text = OpenFileDialogx.FileName;
        }
        private void button4_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox9.Text = OpenFileDialogx.FileName;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox10.Text = OpenFileDialogx.FileName;
        }
        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox11.Text = OpenFileDialogx.FileName;

        }
        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "XML fichier .xml|*.xml";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox13.Text = OpenFileDialogx.FileName;
        }
        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox14.Text = OpenFileDialogx.FileName;
        }


        private void button13_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox15.Text = OpenFileDialogx.FileName;
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox16.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            
            try
            {
                int time1x = System.Environment.TickCount;
                Index_Files cIndex = new Index_Files(textBox16.Text, "PrefaceNP", textBox15.Text, "D:\\ptw\\divi\\ACT1.xlsx");
                textBox20.AppendText("==> Start Création des fichiers Index" + System.Environment.NewLine);
                if (radioButton1.Checked)
                {
                    cIndex.CreateFiles(true);
                }
                else
                {
                    cIndex.CreateFiles(false);
                }
                int time2 = System.Environment.TickCount;

                int times = time2 - time1x;
                string timcIndex = Convert.ToString(Convert.ToDecimal(times) / 1000);
                int hours = 0;
                int minuit = times / 60;
                int second = times - minuit * 60 - hours * 3600;
                timcIndex = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Création des fichiers Index : " + timcIndex + " secondes" + System.Environment.NewLine);
                MessageBox.Show("Création des fichiers Index : " + timcIndex + " secondes");
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString()+ System.Environment.NewLine);
            }
           

         
        }

        //Repertoire Destitation pour les fichier subdiviser et styler
        private void button37_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox6.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }
        //Repertoire source pour les fichier à fusionner
        private void button31_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox4.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }
        //Repertoire Destitation pour les fichier fusionner
        private void button32_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox5.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }
        //choiser les fichiers ptw pour fusionner
        private void button33_Click(object sender, EventArgs e)
        {
            if (checkBox6.Checked == false && checkBox7.Checked == false && checkBox8.Checked == false && checkBox9.Checked == false && checkBox10.Checked == false && checkBox11.Checked == false && checkBox3.Checked == false)
            {
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;
                checkBox11.Checked = true;
                checkBox3.Checked = true;
            }
            else
            {
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox3.Checked = false;
            }
        }
        //supprimer style pour fusionner
        private void button34_Click(object sender, EventArgs e)
        {
            if (checkBox12.Checked == false && checkBox13.Checked == false && checkBox14.Checked == false && checkBox15.Checked == false && checkBox16.Checked == false && checkBox17.Checked == false && checkBox18.Checked == false)
            {
                checkBox12.Checked = true;
                checkBox13.Checked = true;
                checkBox14.Checked = true;
                checkBox15.Checked = true;
                checkBox16.Checked = true;
                checkBox17.Checked = true;
                checkBox18.Checked = true;
            }
            else
            {
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
                checkBox18.Checked = false;
            }
        }
        //preparation nota-pme

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true) 
            {
                textBox3.Enabled = true;
                button28.Enabled = true;
            }
            if (checkBox4.Checked == false)
            {
                textBox3.Enabled = false;
                button28.Enabled = false;
            }
        }


        private void SvparDefault_Click(object sender, EventArgs e)
        {
            string filePath = "D:\\ptw\\pathinfo.ini";
            IniFile iniFile = new IniFile(filePath);
            iniFile.WriteInivalue("dossier", "pathsource", textBox1.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathxml", textBox2.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathdestinationdivi", textBox3.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathdestinationstyle", textBox6.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathdestinationfusion", textBox5.Text.ToString());
            iniFile.WriteInivalue("dossier", "sourcestyle", textBox7.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathSourceFusion", textBox4.Text.ToString());

            //diviser
            iniFile.WriteInivalue("dossier", "sourcedivi", textBox9.Text.ToString());
            iniFile.WriteInivalue("dossier", "styledivi", textBox10.Text.ToString());
            iniFile.WriteInivalue("dossier", "sourceprefaceNP", textBox11.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathprefaceNP", textBox12.Text.ToString());

            iniFile.WriteInivalue("dossier", "stylefusion", textBox13.Text.ToString());
            iniFile.WriteInivalue("dossier", "styletest", textBox8.Text.ToString());

            iniFile.WriteInivalue("dossier", "sourcestyletest", textBox14.Text.ToString());

            iniFile.WriteInivalue("dossier", "sourceindex", textBox15.Text.ToString());
            iniFile.WriteInivalue("dossier", "pathdestinationindex", textBox16.Text.ToString());

        }

        private void button39_Click(object sender, EventArgs e)
        {
            Fussioner_FinalClick(sender, e);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button39_Click_1(object sender, EventArgs e)
        {
            button34_Click_1(sender, e);
         
            MessageBox.Show("Fusion ptw terminée : " + timfusion + " secondes");
        }
        public bool checkptwfile(object sender, EventArgs e)
        {
            string file1 = "D:\\ptw\\Annuel.ptw";
            string file2 = "D:\\ptw\\Admin.ptw";
            string file3 = "D:\\ptw\\Histo.ptw";
            string file4 = "D:\\ptw\\Eval.ptw";
            string file5 = "D:\\ptw\\Decis.ptw";
            string file6 = "D:\\ptw\\Tres.ptw";
            string file7 = "D:\\ptw\\Histo-s.ptw";
            bool flag = true;
            if (File.Exists(file1))
            {
                FileInfo fi = new FileInfo(file1);
                checkBox12.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox12.Checked = true;
            }
            else
            {
                checkBox12.Checked = false;
                flag = false;
            }
            if (File.Exists(file2))
            {
                FileInfo fi = new FileInfo(file2);
                checkBox13.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox13.Checked = true;
            }
            else
            {
                checkBox13.Checked = false;
                flag = false;
            }
            if (File.Exists(file3))
            {
                FileInfo fi = new FileInfo(file3);
                checkBox14.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox14.Checked = true;
            }
            else
            {
                checkBox14.Checked = false;
                flag = false;
            }
            if (File.Exists(file4))
            {
                FileInfo fi = new FileInfo(file4);
                checkBox15.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox15.Checked = true;
            }
            else
            {
                flag = false;
                checkBox15.Checked = false;
            }
            if (File.Exists(file5))
            {
                FileInfo fi = new FileInfo(file5);
                checkBox16.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox16.Checked = true;
            }
            else
            {
                flag = false;
                checkBox16.Checked = false;
            }
            if (File.Exists(file6))
            {
                FileInfo fi = new FileInfo(file6);
                checkBox17.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox17.Checked = true;
            }
            else
            {
                flag = false;
                checkBox17.Checked = false;
            }
            if (File.Exists(file7))
            {
                FileInfo fi = new FileInfo(file7);
                checkBox18.Text = fi.LastWriteTime.ToLocalTime().ToString();
                checkBox18.Checked = true;
            }
            else
            {
                flag = false;
                checkBox18.Checked = false;
            }
            if (flag)
            {
                Fussioner_FinalClick(sender, e);
            }
            else
            {
                MessageBox.Show("File missing");
            }
            return flag;
        }
        private void button34_Click_1(object sender, EventArgs e)
        {
            checkptwfile(sender, e);
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
                textBox9.Text = @"D:\ptw\prefaceNP.xlsx";
            if (radioButton3.Checked == true)
                textBox9.Text = @"D:\ptw\preface.xlsx";
            if (radioButton4.Checked == true)
                textBox9.Text = @"D:\ptw\Histo.ptw";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
                textBox9.Text = @"D:\ptw\prefaceNP.xlsx";
            if (radioButton3.Checked == true)
                textBox9.Text = @"D:\ptw\preface.xlsx";
            if (radioButton4.Checked == true)
                textBox9.Text = @"D:\ptw\Histo.ptw";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
                textBox9.Text = @"D:\ptw\prefaceNP.xlsx";
            if (radioButton3.Checked == true)
                textBox9.Text = @"D:\ptw\preface.xlsx";
            if (radioButton4.Checked == true)
                textBox9.Text = @"D:\ptw\Histo.ptw";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                string timex = System.DateTime.Now.Day + "." + System.DateTime.Now.Month + "." + System.DateTime.Now.Year + " à " + System.DateTime.Now.Hour + "." + System.DateTime.Now.Minute + "." + System.DateTime.Now.Second;
                textBox20.AppendText("==> Début " + timex);
                textBox20.AppendText("==> Démarrage des processus... " + System.Environment.NewLine);
                int time1 = System.Environment.TickCount;
                int timestyle = 0;


                string Protection = "";
                string formatzro = "";
                string formatprefacenpH = "";
                string formatprefacenpC = "";
                int indexfiles = 0;
                bool flag = true;
                //Alex: creation du style
                if (checkBox34.Checked)
                {
                    Xmlecriretout_Click(sender, e);
                    int timex1 = System.Environment.TickCount;
                    timestyle = timex1 - time1;
                }
                if (checkBox26.Checked)
                {
                    flag = checkptwfile(sender, e);

                }
                else
                {
                    flag = true;
                }
                if (flag)
                {
                    if (checkBox27.Checked)
                    {
                        leger_Click(sender, e);
                    }
                    if (checkBox31.Checked)
                    {
                        Protection = button20tout_Click(sender, e);
                    }
                    if (checkBox33.Checked)
                    {
                        formatzro = button19_Click_tout(sender, e);
                    }
                    if (checkBox28.Checked)
                    {
                        try
                        {
                            textBox20.AppendText("==> Start Découpage des fichiers" + System.Environment.NewLine);
                            if (checkBox19.Checked == true)
                            {
                                Diviser_Click(sender, e);//historique
                            }
                            if (checkBox21.Checked == true)
                            {
                                DiviserHistS(sender, e);//historique-s
                            }
                            if (checkBox22.Checked == true)
                            {
                                diviserAnnuel(sender, e);//comptes Annuels
                            }
                            if (checkBox23.Checked == true)
                            {
                                diviserSynthese(sender, e);//SynthèseValorisations
                            }
                            button19_Click(sender, e);
                            textBox20.AppendText("Découpage Histo.ptw,    " + timdiviser + " secondes" + Environment.NewLine +
                                         "Découpage Histo-s.ptw, " + timdiviserHistoS + " secondes" + Environment.NewLine +
                                         "Découpage Annuel.ptw, " + timdiviserAnnuel + " secondes" + Environment.NewLine +
                                         "Découpage Eval.ptw, " + timdiviserSynthese + " secondes" + Environment.NewLine);
                        }
                        catch (Exception ex)
                        {
                            textBox20.AppendText(ex + System.Environment.NewLine);
                        }
                    }
                    if (checkBox32.Checked)
                    {
                        int timex1 = System.Environment.TickCount;
                        Index_Files cIndex = new Index_Files(textBox16.Text, "PrefaceNP", textBox15.Text, "D:\\ptw\\divi\\ACT1.xlsx");
                        textBox20.AppendText("==> Start Création des fichiers Index" + System.Environment.NewLine);
                        cIndex.CreateFiles(false);


                        int timex2 = System.Environment.TickCount;
                        indexfiles = timex2 - timex1;
                        string timcIndex = Convert.ToString(Convert.ToDecimal(indexfiles) / 1000);

                        int hoursX = indexfiles / 3600;
                        int minuitX = indexfiles / 60 - hoursX * 60;
                        int secondX = indexfiles - minuitX * 60 - hoursX * 3600;
                        timcIndex = hoursX + " heures " + minuitX + " minutes " + secondX;
                        textBox20.AppendText("Création des fichiers Index : " + timcIndex + " secondes" + System.Environment.NewLine);
                    }
                    if (checkBox29.Checked)
                    {
                        formatprefacenpH = button23_Clicktout(sender, e);
                    }
                    if (checkBox30.Checked)
                    {
                        formatprefacenpC = button26_Clicktout(sender, e);
                    }
                    //Diviser_Click(sender, e);
                    int time2 = System.Environment.TickCount;

                    int times = (time2 - time1) / 1000;
                    int hours = times / 3600;
                    int minuit = times / 60 - hours * 60;
                    int second = times - minuit * 60 - hours * 3600;
                    timtotal = hours + " heures " + minuit + " minutes " + second;
                    //timtotal = Convert.ToString(Convert.ToDecimal(times) / 1000);
                    string showtext = "Temps total ----------------------------" + timtotal + Environment.NewLine + Environment.NewLine;
                    if (checkBox34.Checked)
                    {
                        showtext = showtext + "Création des fichiers de style              " + timestyle / 1000 + " secondes" + Environment.NewLine;
                    }
                    if (checkBox26.Checked)
                    {
                        showtext = showtext + "Fusion des fichiers *.ptw              " + timfusion + Environment.NewLine;
                    }
                    if (checkBox27.Checked)
                    {
                        showtext = showtext + Environment.NewLine + "Création de PrefaceNP.xlsx                 " + timleger + Environment.NewLine;
                    }
                    if (checkBox31.Checked)
                    {
                        showtext = showtext + "Protection des cellules OK : " + Protection + " s" + Environment.NewLine;
                    }
                    if (checkBox33.Checked)
                    {
                        showtext = showtext + "Raz PrefaceNP OK. Le fichier est sauvé dans D:\\ptw\\notepme. Temps : " + formatzro + " s";
                    }
                    if (checkBox28.Checked)
                    {
                        showtext = showtext + "Création des sous-fichiers pour Histo.ptw      " + timdiviser + Environment.NewLine +
                                    "Création des sous-fichiers pour Histo-s.ptw     " + timdiviserHistoS + Environment.NewLine +
                                    "Création des sous-fichiers pour Annuel.ptw     " + timdiviserAnnuel + Environment.NewLine +
                                    "Création des sous-fichiers pour Eval.ptw     " + timdiviserSynthese + Environment.NewLine;
                    }
                    if (checkBox29.Checked)
                    {
                        showtext = showtext + "Preface style format History OK : " + formatprefacenpH + " s" + Environment.NewLine;
                    }
                    if (checkBox30.Checked)
                    {
                        showtext = showtext + "Preface style format Comptes annuel OK : " + formatprefacenpC + " s" + Environment.NewLine;
                    }

                    if (checkBox32.Checked)
                    {
                        showtext = showtext + "Création des fichiers Index : " + indexfiles / 1000 + " secondes" + Environment.NewLine;
                    }

                    textBox20.AppendText(System.Environment.NewLine + "Terminé. Temps total : " + timtotal + System.Environment.NewLine);
                    string timexx = System.DateTime.Now.Day + "." + System.DateTime.Now.Month + "." + System.DateTime.Now.Year + " à " + System.DateTime.Now.Hour + "." + System.DateTime.Now.Minute + "." + System.DateTime.Now.Second;
                    textBox20.AppendText("==> Fin " + timexx);
                    MessageBox.Show("==> Début " + timex + System.Environment.NewLine + showtext + System.Environment.NewLine + "==> Fin " + timexx);
                }
            }
            else
            {
                string timex = System.DateTime.Now.Day + "." + System.DateTime.Now.Month + "." + System.DateTime.Now.Year + " à " + System.DateTime.Now.Hour + "." + System.DateTime.Now.Minute + "." + System.DateTime.Now.Second;
                textBox20.AppendText("==> Début " + timex);
                textBox20.AppendText("==> Démarrage des processus... " + System.Environment.NewLine);
                int time1 = System.Environment.TickCount;
                if (checkBox26.Checked)
                {
                    button34_Click_1(sender, e);

                }
                if (checkBox27.Checked)
                {
                    simply_leger_Click(sender, e);
                    sheetnamechange_Click(sender, e);
                }
                if (checkBox31.Checked && checkBox31.Checked)
                {
                    flagsimplypastout = false;
                    LockStateFiles(sender, e);
                }
                if (checkBox28.Checked)
                {
                    createhissimply(sender, e);
                }
                string timexx = System.DateTime.Now.Day + "." + System.DateTime.Now.Month + "." + System.DateTime.Now.Year + " à " + System.DateTime.Now.Hour + "." + System.DateTime.Now.Minute + "." + System.DateTime.Now.Second;
                int timex2 = System.Environment.TickCount;
               long indexfiles = timex2 - time1;
                string timcIndex = Convert.ToString(Convert.ToDecimal(indexfiles) / 1000);

                long hoursX = indexfiles / 3600000;
                long minuitX = indexfiles / 60 - hoursX * 60;
                long secondX = indexfiles - minuitX * 60 - hoursX * 3600;
                timcIndex = hoursX + " heures " + minuitX + " minutes " + secondX;
                textBox20.AppendText(System.Environment.NewLine + "Terminé. Temps total : " + timtotal + System.Environment.NewLine);
                MessageBox.Show("==> Début " + timex + System.Environment.NewLine + timcIndex + System.Environment.NewLine + "==> Fin " + timexx);
            }
        }

#endregion

        #region Histo-s

        //////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////Hist-s.ptw///////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////

        //Histo Simplifier

        private void renommer()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            //xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            //Excel.Worksheet sheetTypologie = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Typologie IFRS");
            //sheetTypologie.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        //nom de colonne solid a modifier
        private void insertionHistoS(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;

            Excel.Range rangex1 = xlWorkSheet.Cells[1, 4] as Excel.Range;

            Excel.Range rangex2 = xlWorkSheet.Cells[1, 5] as Excel.Range;

            Excel.Range rangex3 = xlWorkSheet.Cells[1, 6] as Excel.Range;


            //rangex1c.EntireColumn.Copy(misValue);
            //rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangex1c.EntireColumn.Copy(misValue);
            //rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangex1c.EntireColumn.Copy(misValue);
            //rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangex1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangex2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangex3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            Excel.Range rangex1c = xlWorkSheet.UsedRange.get_Range("EU1", "EV1") as Excel.Range;
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("D1", "E1").EntireColumn);
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("G1", "H1").EntireColumn);
            rangex1c.EntireColumn.Copy(xlWorkSheet.UsedRange.get_Range("J1", "K1").EntireColumn);


            Excel.Worksheet xlWorkSheet2 = xlWorkBook.Worksheets["Hist.Refer-s"] as Excel.Worksheet;
            //Excel.Range rangeC = xlWorkSheet2.Cells[1, 4] as Excel.Range;

            //Excel.Range rangeD = xlWorkSheet2.Cells[1, 5] as Excel.Range;

            //Excel.Range rangeE = xlWorkSheet2.Cells[1, 6] as Excel.Range;

            //rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            //rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            Excel.Range rangeCc = xlWorkSheet2.Cells[1, 3] as Excel.Range;
            Excel.Range rangeCc2 = xlWorkSheet2.Cells[1, 4] as Excel.Range;
            Excel.Range rangeCc3 = xlWorkSheet2.Cells[1, 5] as Excel.Range;
            Excel.Range rangeDc = xlWorkSheet2.Cells[1, 6] as Excel.Range;
            Excel.Range rangeDc2 = xlWorkSheet2.Cells[1, 7] as Excel.Range;
            Excel.Range rangeDc3 = xlWorkSheet2.Cells[1, 8] as Excel.Range;
            Excel.Range rangeEc = xlWorkSheet2.Cells[1, 9] as Excel.Range;
            Excel.Range rangeEc2 = xlWorkSheet2.Cells[1, 10] as Excel.Range;
            Excel.Range rangeEc3 = xlWorkSheet2.Cells[1, 11] as Excel.Range;

            rangeCc.EntireColumn.Copy(rangeCc2.EntireColumn);
            rangeCc.EntireColumn.Copy(rangeCc3.EntireColumn);

            rangeCc.EntireColumn.Copy(rangeDc.EntireColumn);
            rangeCc.EntireColumn.Copy(rangeDc2.EntireColumn);
            rangeCc.EntireColumn.Copy(rangeDc3.EntireColumn);

            rangeCc.EntireColumn.Copy(rangeEc.EntireColumn);
            rangeCc.EntireColumn.Copy(rangeEc2.EntireColumn);
            rangeCc.EntireColumn.Copy(rangeEc3.EntireColumn);

            //Excel.Worksheet WorkSheetPreface = xlWorkBook.Worksheets["Hist.Preface"] as Excel.Worksheet;


            Excel.Range rangex1cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 4] as Excel.Range;
            Excel.Range rangex1cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 5] as Excel.Range;
            Excel.Range rangex2cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 7] as Excel.Range;
            Excel.Range rangex2cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 8] as Excel.Range;
            Excel.Range rangex3cx1 = xlWorkSheet.Cells[range.Rows.Count - 1, 10] as Excel.Range;
            Excel.Range rangex3cx2 = xlWorkSheet.Cells[range.Rows.Count - 1, 11] as Excel.Range;

            rangex1cx1.Value2 = "";
            rangex1cx2.Value2 = "";
            rangex2cx1.Value2 = "";
            rangex2cx2.Value2 = "";
            rangex3cx1.Value2 = "";
            rangex3cx2.Value2 = "";




            //tester EE pour Histo.refer//et parcourir historique

            Excel.Range rangeRefer = xlWorkSheet2.UsedRange;
            Excel.Range rangeHistorique = xlWorkSheet.UsedRange;
            //petite corr
            object[,] valuesRefer = (object[,])rangeRefer.Value2;
            object[,] valuesHistorique = (object[,])rangeHistorique.Value2;
            int rowCnt = 0;
            int rowHistoCnt = 0;
            string nomCol = "";

            for (rowCnt = 1; rowCnt <= rangeRefer.Rows.Count; rowCnt++)
            {
                string valuecellabs = Convert.ToString(valuesRefer[rowCnt, 1]);
                if (valuecellabs != "" && valuecellabs != "D" && valuecellabs != "D1" && valuecellabs != "d")
                {
                    nomCol = valuecellabs;
                    for (rowHistoCnt = 1; rowHistoCnt <= rangeHistorique.Rows.Count; rowHistoCnt++)
                    {
                        string valuecellHisto = Convert.ToString(valuesHistorique[rowHistoCnt, 2]);
                        if (valuecellHisto == nomCol)
                        {
                            Excel.Range cellcopie = xlWorkSheet.Cells[rowHistoCnt, 3] as Excel.Range;
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 4]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 5]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 6]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 7]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 8]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 9]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 10]);
                            cellcopie.Copy(xlWorkSheet.Cells[rowHistoCnt, 11]);
                        }
                    }
                }
            }




            ///////////////////////////////Parcourir tous les cellule ergodiaue////////////////////////////////
            //string str;
            //int rCnt = 0;
            //int cCnt = 0;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //    {
            //        str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString();
            //        MessageBox.Show(str);
            //    }
            //}
            ///////////////////////////////fermer EXCEL automatiquement apres modification?//////////////////////
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            //MessageBox.Show("jobs done!");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void supprimercolhistoS(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////////////400000//////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row400000 = 0;
            cCnt = range.Columns.Count;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row400000 = cf.FindCodedRow("400000", range);



            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "400000"))
            //    {
            //        row400000 = rCnt;
            //        break;
            //    }
            //}

            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row400000, col]);
                if (Regex.Equals(value, "-1"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row400000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void consigneProtegerHistoS()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            int rowcount = xlWorkSheet.UsedRange.Rows.Count;
            object[,] values = (object[,])range.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            rCnt = range.Rows.Count;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col = cf.FindCodedColumn("15000", range);

            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "15000"))
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}

            //Routine pour modifier col XXXXX marquer ligne proteger -1
            for (int i = 1; i < rowcount - 5; i++)
            {
                if ((xlWorkSheet.Cells[i, 3] as Excel.Range).Locked.ToString() != "True")
                    (xlWorkSheet.Cells[i, col] as Excel.Range).Value2 = "-1";
            }


            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void HistoprefaceHistoS(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            //prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
          

            Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets["Hist.Preface-s"] as Excel.Worksheet;

            Excel.Range rangeinsert1 = xlWorkSheet.UsedRange.get_Range("D1", "E1") as Excel.Range;
            rangeinsert1.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();
            Excel.Range rangeinsert2 = xlWorkSheet.UsedRange.get_Range("H1", "I1") as Excel.Range;
            rangeinsert2.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();
            Excel.Range rangeinsert3 = xlWorkSheet.UsedRange.get_Range("L1", "M1") as Excel.Range;
            rangeinsert3.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            xlWorkBook.Save();
            releaseObject(rangeinsert1);
            releaseObject(rangeinsert2);
            releaseObject(rangeinsert3);



            Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
            Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
            Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
            rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
            rangeReplace1.EntireColumn.Replace("'Historique-s'!C", "'Historique-s'!E", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace1.EntireColumn.Replace("Hist.Refer-s!A", "Hist.Refer-s!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace1.EntireColumn.Copy(rangeOrigin1.EntireColumn);
            rangeReplace1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            releaseObject(rangeOrigin1);
            releaseObject(rangeMiddle1);
            releaseObject(rangeReplace1);

            //2
            Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
            Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
            Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
            rangeReplace2.EntireColumn.Replace("'Historique-s'!F", "'Historique-s'!H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace2.EntireColumn.Replace("Hist.Refer-s!D", "Hist.Refer-s!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace2.EntireColumn.Copy(rangeOrigin2.EntireColumn);
            rangeReplace2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            releaseObject(rangeOrigin2);
            releaseObject(rangeMiddle2);
            releaseObject(rangeReplace2);

            //3
            Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
            Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
            Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
            rangeReplace3.EntireColumn.Replace("'Historique-s'!I", "'Historique-s'!K", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace3.EntireColumn.Replace("Hist.Refer-s!G", "Hist.Refer-s!I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            rangeReplace3.EntireColumn.Copy(rangeOrigin3.EntireColumn);
            rangeReplace3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
            releaseObject(rangeOrigin3);
            releaseObject(rangeMiddle3);
            releaseObject(rangeReplace3);

            ////1
            //Excel.Range rangeOrigin1 = xlWorkSheet.Cells[1, 6] as Excel.Range;
            //Excel.Range rangeMiddle1 = xlWorkSheet.Cells[1, 5] as Excel.Range;
            //Excel.Range rangeReplace1 = xlWorkSheet.Cells[1, 4] as Excel.Range;
            //rangeOrigin1.EntireColumn.Copy(rangeReplace1.EntireColumn);
            //rangeReplace1.EntireColumn.Replace("Historique-s!A", "Historique-s!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace1.EntireColumn.Replace("Hist.Refer-s!A", "Hist.Refer-s!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace1.EntireColumn.Copy(rangeOrigin1.EntireColumn);
            //rangeReplace1.EntireColumn.Copy(rangeMiddle1.EntireColumn);
            //releaseObject(rangeOrigin1);
            //releaseObject(rangeMiddle1);
            //releaseObject(rangeReplace1);

            ////2
            //Excel.Range rangeOrigin2 = xlWorkSheet.Cells[1, 10] as Excel.Range;
            //Excel.Range rangeMiddle2 = xlWorkSheet.Cells[1, 9] as Excel.Range;
            //Excel.Range rangeReplace2 = xlWorkSheet.Cells[1, 8] as Excel.Range;
            //rangeOrigin2.EntireColumn.Copy(rangeReplace2.EntireColumn);
            //rangeReplace2.EntireColumn.Replace("Historique-s!D", "Historique-s!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace2.EntireColumn.Replace("Hist.Refer-s!D", "Hist.Refer-s!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace2.EntireColumn.Copy(rangeOrigin2.EntireColumn);
            //rangeReplace2.EntireColumn.Copy(rangeMiddle2.EntireColumn);
            //releaseObject(rangeOrigin2);
            //releaseObject(rangeMiddle2);
            //releaseObject(rangeReplace2);

            ////3
            //Excel.Range rangeOrigin3 = xlWorkSheet.Cells[1, 14] as Excel.Range;
            //Excel.Range rangeMiddle3 = xlWorkSheet.Cells[1, 13] as Excel.Range;
            //Excel.Range rangeReplace3 = xlWorkSheet.Cells[1, 12] as Excel.Range;
            //rangeOrigin3.EntireColumn.Copy(rangeReplace3.EntireColumn);
            //rangeReplace3.EntireColumn.Replace("Historique-s!G", "Historique-s!I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace3.EntireColumn.Replace("Hist.Refer-s!G", "Hist.Refer-s!I", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeReplace3.EntireColumn.Copy(rangeOrigin3.EntireColumn);
            //rangeReplace3.EntireColumn.Copy(rangeMiddle3.EntireColumn);
            //releaseObject(rangeOrigin3);
            //releaseObject(rangeMiddle3);
            //releaseObject(rangeReplace3);

            Excel.Range rangeRef = xlWorkSheet.UsedRange;

            object[,] values = (object[,])rangeRef.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int Row500000 = 0;
            cCnt = rangeRef.Columns.Count;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            Row500000 = cf.FindCodedRow("500000", rangeRef);

            //for (rCnt = 1; rCnt <= rangeRef.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "500000"))
            //    {
            //        Row500000 = rCnt;
            //        break;
            //    }
            //}
            //Excel.Range rangeXLReplace1 = xlWorkSheet.Cells[Row500000, 8] as Excel.Range;
            //Excel.Range rangeXLC11 = xlWorkSheet.Cells[Row500000, 9] as Excel.Range;
            //Excel.Range rangeXLC12 = xlWorkSheet.Cells[Row500000, 10] as Excel.Range;
            ////eviter bug vsto
            //xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[Row500000, 8], xlWorkSheet.Cells[Row500000, 9]).Replace("Historique-s!A", "Historique-s!C", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeXLReplace1.Copy(rangeXLC11);
            //rangeXLReplace1.Copy(rangeXLC12);


            //Excel.Range rangeXLReplace2 = xlWorkSheet.Cells[Row500000, 12] as Excel.Range;
            //Excel.Range rangeXLC21 = xlWorkSheet.Cells[Row500000, 13] as Excel.Range;
            //Excel.Range rangeXLC22 = xlWorkSheet.Cells[Row500000, 14] as Excel.Range;
            ////eviter bug vsto
            //xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[Row500000, 12], xlWorkSheet.Cells[Row500000, 13]).Replace("Historique-s!D", "Historique-s!F", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            //rangeXLReplace2.Copy(rangeXLC21);
            //rangeXLReplace2.Copy(rangeXLC22);



            //xlWorkBook.SaveCopyAs(prefaceNP);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        public void legerHistS(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;

            //2003-2010!!!!!!!!!!!!!!
            fichierprepare = "D:\\ptw\\preface.xlsx";
            prefaceNP = "D:\\ptw\\prefaceNP.xlsx";

            renommer();

            //button2_Click(sender, e);
            //HistoCalculs();
            //HistoMettreZero_Click(sender, e);
            //HistoRempl_Click(sender, e);
            //HistoAuAvAw_Click(sender, e);
            //colCE_Click(sender, e);//72000
            //supprimerREF_Click(sender, e);

            ////////////Histo.ptw et histo.preface
            insertionHistoS(sender, e);//Inserer les colonnes correctifs
            //Histopreface_Click(sender, e);

            ////////Annuel .ptw
            //AnnuelO_Click(sender, e);
            //ComptesAnnuels_Click(sender, e);
            supprimercolhistoS(sender, e);

            //button5_Click(sender, e);

            //supprimer les onglets
            //Supprimeronglet_Click(sender, e);

            //traitement REF!
            //Historique84000();
            //fonctionRemplacerD1();
            consigneProtegerHistoS();



            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string timlegerHS = Convert.ToString(Convert.ToDecimal(times) / 1000);
            MessageBox.Show(timlegerHS);
        }

        //procedure diviser
        private void supprimermoin2HistS(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////944000//////////////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row400000 = 0;
            cCnt = range.Columns.Count;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row400000 = cf.FindCodedRow("400000", range);

            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "400000"))
            //    {
            //        row400000 = rCnt;
            //        break;
            //    }
            //}

            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row400000, col]);
                if (Regex.Equals(value, "-2"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row400000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }

            range = xlWorkSheet.UsedRange;
            cCnt = range.Columns.Count;
            values = (object[,])range.Value2;
            for (int col = 1; col <= cCnt; col++)
            {
                string value = Convert.ToString(values[row400000, col]);
                if (Regex.Equals(value, "-4"))
                {
                    Excel.Range rangeEffacer = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col], xlWorkSheet.Cells[row400000 - 1, col]) as Excel.Range;
                    rangeEffacer.ClearContents();
                }
            }



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void DiviStylerHistS(object sender, EventArgs e)
        {
            pathnotapme = textBox3.Text;
            pathstylerfinal = textBox6.Text;

            string openfilex = "D:\\ptw\\Histo.xlsx";

            ////////////////open excel///////////////////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Historique-s");
            Excel.Range rangeRefall = xlWorkSheetRef.UsedRange;
            //exception!!!
            xlWorkSheetRef.Cells.ColumnWidth = 20;

            Excel.Range rangeRef = xlWorkSheetRef.Cells[rangeRefall.Rows.Count, 1] as Excel.Range;
            rangeRef.EntireRow.Copy(misValue);
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", xlWorkSheetRef.Cells[rangeRefall.Rows.Count - 1, 1]) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            xlWorkSheetRef.SaveAs("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            //////////////////////////////////////////////////////////////////////////////////
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 = "[" + openfilex.Substring(7, 9) + "]";
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            int col83000 = 0;
            rCnt = range.Rows.Count;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col3000 = cf.FindCodedColumn("3000", range);
            col4000 = cf.FindCodedColumn("4000", range);
            col5000 = cf.FindCodedColumn("5000", range);
            col8000 = cf.FindCodedColumn("8000", range);
            col = cf.FindCodedColumn("10000", range);
            col83000 = cf.FindCodedColumn("83000", range);


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "3000"))
            //    {
            //        col3000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "4000"))
            //    {
            //        col4000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "5000"))
            //    {
            //        col5000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "8000"))
            //    {
            //        col8000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "10000"))
            //    {
            //        col = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "83000"))
            //    {
            //        col83000 = cCnt;
            //        break;
            //    }
            //}
            int fileflag = 0;
            for (int row = 25; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Historique-s");
                    string[] namestable = { "ACT-s1.xlsx", "PAS-s1.xlsx", "CR-s1.xlsx", "CR-s2.xlsx", "ANN-s11.xlsx", "ANN-s12.xlsx", "ANN-s13.xlsx", "ANN-s21.xlsx", "ANN-s31.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire
                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[25, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[25, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[24, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[25, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de row ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range("A25", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Historique-s");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);
                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Historique-s");
                    Excel.Range rangeremplace = xlWorkSheetremplace.UsedRange;
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas
                    ////////delete col8000 "-2"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    {
                        string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                        if (Regex.Equals(valuedel, "-2"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowdel--;
                        }
                    }
                    ///////////////row hide "-5"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }
                    ///////////////Hide -1 pour col 83000/////////////////////////////////////////////
                    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                    //    if (Regex.Equals(valuedel, "-1"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col83000] as Excel.Range;
                    //        rangeDely.EntireRow.Hidden = true;
                    //    }
                    //}
                    /////////////////////////////////////////////////////////////////////////////////
                    object[,] valuesNX = (object[,])rangeremplace.Value2;
                    //string valueNX = Convert.ToString(valuesNX[row, col]);
                    for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                        {
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row3000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row4000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row5000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }

                    ////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    if (checkBox20.Checked == true)
                    {
                        fileAstyler = divisavenom;
                        Xmllire_Click(sender, e);
                    }

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = 25;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void DiviserHistS(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            fichierprepare = "D:\\ptw\\prefaceNP.xlsx";// textBox9.Text;
            prefaceNP = "D:\\ptw\\Histo.xlsx";


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //ne afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;

            //Hist.Refer
            Excel.Worksheet sheetHistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");
            Excel.Range rangeHistRefer = sheetHistRefer.UsedRange;
            rangeHistRefer.Copy(misValue);
            rangeHistRefer.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            ////Hist.Refer mise à zero// dégalage?
            object[,] valuesRefer = (object[,])sheetHistRefer.UsedRange.Value2;

            for (int rowCnt = 1; rowCnt <= rangeHistRefer.Rows.Count-1; rowCnt++)//sauf derniere ligne..
            {
                string valuecellabs = Convert.ToString(valuesRefer[rowCnt, 1]);
                if (valuecellabs != "")
                {
                    Excel.Range referZero = sheetHistRefer.UsedRange.get_Range(sheetHistRefer.UsedRange.Cells[rowCnt, 3], sheetHistRefer.UsedRange.Cells[rowCnt, 11]) as Excel.Range;
                    //referZero.Copy();
                    referZero.Value2 = 0;
                    //D1 D2 D3  =""
                    if (valuecellabs == "D" || valuecellabs == "D1" || valuecellabs == "d")
                    {
                        referZero.Formula = "=\"\"";
                    }
                }
            }




            //suppression des onglets
            Excel.Worksheet Historique = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet HistPrefac = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet HistCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Worksheet HistLangues = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues");
            Excel.Worksheet HistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");


            //Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-s");
            Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-s");
            //Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-s");

            Excel.Worksheet ComptesannuelRefssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Worksheet Comptesannuelssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            Excel.Worksheet Paramimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            Excel.Worksheet Paramgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            Excel.Worksheet AdminLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            Excel.Worksheet AdminServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            Excel.Worksheet Tsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            Excel.Worksheet ParamSavsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            Excel.Worksheet Macrossheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            Excel.Worksheet Vsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            Excel.Worksheet Mosaiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            Excel.Worksheet GraphiquesSRsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            Excel.Worksheet Graphimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            Excel.Worksheet Dontdeletesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            Excel.Worksheet Finsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            Excel.Worksheet ChoixMethodessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            Excel.Worksheet Noterecapitulativesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            Excel.Worksheet SyntheseValorisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Worksheet DefinitionsArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            Excel.Worksheet RappelRetraitementssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            Excel.Worksheet RisqueEntreprisesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            Excel.Worksheet ChoixTauxParamsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            Excel.Worksheet TauxParamArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            Excel.Worksheet CorrectifsSIGBilansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            Excel.Worksheet APNNEsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            Excel.Worksheet FiscaliteDiffereesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            Excel.Worksheet PatrimonialAncAnccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            Excel.Worksheet FondsDeCommercesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            Excel.Worksheet Goodwillsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            Excel.Worksheet AutresCapitalisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            Excel.Worksheet Multiplessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            Excel.Worksheet MethodesMixtessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            Excel.Worksheet TransactionsComparablessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            Excel.Worksheet GordonShapiroBatessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            Excel.Worksheet CalculFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            Excel.Worksheet DiscountedFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            Excel.Worksheet CmpcWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            Excel.Worksheet CmpcWaccArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            Excel.Worksheet ModuleWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            Excel.Worksheet CCEFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            Excel.Worksheet TriRentabiliteProjetsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            Excel.Worksheet TourDeTableSynthesesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            Excel.Worksheet EvalLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            Excel.Worksheet Controlessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            Excel.Worksheet EvalServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            Excel.Worksheet Composantessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            Excel.Worksheet Jsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            Excel.Worksheet Factgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            Excel.Worksheet Lsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            Excel.Worksheet Msheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            Excel.Worksheet Tresoreriesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            Excel.Worksheet ABsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            Excel.Worksheet Paramtresorsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            Excel.Worksheet Saisonnalitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            Excel.Worksheet Zsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
            Excel.Worksheet sheetCA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CA");
            Excel.Worksheet sheetInvestissements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investissements");
            Excel.Worksheet sheetCpteresultat = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Cpte Résultat");
            Excel.Worksheet sheetFinancements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Financements");
            Excel.Worksheet sheetbfr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("BFR");
            Excel.Worksheet sheetbilan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Bilan");
            Excel.Worksheet sheetcontrole2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles (2)");
            Excel.Worksheet sheetmultiple = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiple");
            Excel.Worksheet sheetvalo = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Valo et ouverture du capital");
            Excel.Worksheet sheetplan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Plan de financement");
            Excel.Worksheet sheetsynthese = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Synthèse SIG et SR");

            sheetCA.Delete();
            sheetInvestissements.Delete();
            sheetCpteresultat.Delete();
            sheetFinancements.Delete();
            sheetbfr.Delete();
            sheetbilan.Delete();
            sheetcontrole2.Delete();
            sheetmultiple.Delete();
            sheetvalo.Delete();
            sheetplan.Delete();
            sheetsynthese.Delete();
            model.Delete();
            Historique.Delete();
            HistPrefac.Delete();
            HistCalculs.Delete();
            HistLangues.Delete();
            HistRefer.Delete();

            HistPrefacsheet.Delete();
            HistCalculssheet.Delete();

            ComptesannuelRefssheet.Delete();
            Comptesannuelssheet.Delete();
            Osheet.Delete();
            Identitesheet.Delete();
            Paramimprsheet.Delete();
            Psheet.Delete();
            Paramgenerauxsheet.Delete();
            AdminLanguessheet.Delete();
            AdminServicesheet.Delete();
            Tsheet.Delete();
            ParamSavsheet.Delete();
            Macrossheet.Delete();
            Vsheet.Delete();
            Mosaiquesheet.Delete();
            GraphiquesSRsheet.Delete();
            Graphimprsheet.Delete();
            Dontdeletesheet.Delete();
            Finsheet.Delete();
            ChoixMethodessheet.Delete();
            Noterecapitulativesheet.Delete();
            SyntheseValorisationssheet.Delete();
            DefinitionsArrierePlansheet.Delete();
            RappelRetraitementssheet.Delete();
            RisqueEntreprisesheet.Delete();
            ChoixTauxParamsheet.Delete();
            TauxParamArrierePlansheet.Delete();
            CorrectifsSIGBilansheet.Delete();
            APNNEsheet.Delete();
            FiscaliteDiffereesheet.Delete();
            PatrimonialAncAnccsheet.Delete();
            FondsDeCommercesheet.Delete();
            Goodwillsheet.Delete();
            AutresCapitalisationssheet.Delete();
            Multiplessheet.Delete();
            MethodesMixtessheet.Delete();
            TransactionsComparablessheet.Delete();
            GordonShapiroBatessheet.Delete();
            CalculFCFsheet.Delete();
            DiscountedFCFsheet.Delete();
            CmpcWaccsheet.Delete();
            CmpcWaccArrierePlansheet.Delete();
            ModuleWaccsheet.Delete();
            CCEFsheet.Delete();
            TriRentabiliteProjetsheet.Delete();
            TourDeTableSynthesesheet.Delete();
            EvalLanguessheet.Delete();
            Controlessheet.Delete();
            EvalServicesheet.Delete();
            Composantessheet.Delete();
            Jsheet.Delete();
            Factgenerauxsheet.Delete();
            Lsheet.Delete();
            Msheet.Delete();
            Tresoreriesheet.Delete();
            ABsheet.Delete();
            Paramtresorsheet.Delete();
            Saisonnalitesheet.Delete();
            Zsheet.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            supprimermoin2HistS(sender, e);
            //subdiviser histo-s
            DiviStylerHistS(sender, e);

            int time2 = System.Environment.TickCount;
            int times = (time2 - time1) / 1000;
            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;

            timdiviserHistoS = hours + " heures " + minuit + " minutes " + second;
            //timdiviserHistoS = Convert.ToString(Convert.ToDecimal(times) / 1000);

            //MessageBox.Show("jobs done " + tim + " seconds used");
        }
#endregion

        #region Annuel
        //////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////Annuel.ptw///////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////

        //Annuel      Comptes annuels

        private void supprimermoin2Annuel(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //demasquer tout
            range.EntireColumn.Hidden = false;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////944000//////////////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row1012000 = 0;

            cCnt = range.Columns.Count;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row1012000 = cf.FindCodedRow("1012000", range);

            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "1012000"))
            //    {
            //        row1012000 = rCnt;
            //        break;
            //    }
            //}

            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row1012000, col]);
                if (Regex.Equals(value, "-2"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row1012000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }

            range = xlWorkSheet.UsedRange;
            cCnt = range.Columns.Count;
            values = (object[,])range.Value2;
            for (int col = 1; col <= cCnt; col++)
            {
                string value = Convert.ToString(values[row1012000, col]);
                if (Regex.Equals(value, "-4"))
                {
                    Excel.Range rangeEffacer = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col], xlWorkSheet.Cells[row1012000 - 1, col]) as Excel.Range;
                    rangeEffacer.ClearContents();
                }
            }



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void DiviStylerAnnuel(object sender, EventArgs e)
        {
            pathnotapme = textBox3.Text;
            pathstylerfinal = textBox6.Text;

            string openfilex = "D:\\ptw\\Histo.xlsx";
            Thread.Sleep(3000);
            ////////////////open excel///////////////////////////////////////
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Comptes annuels");
            Excel.Range rangeRefall = xlWorkSheetRef.UsedRange;
            //exception!!!
            xlWorkSheetRef.Cells.ColumnWidth = 20;

            Excel.Range rangeRef = xlWorkSheetRef.Cells[rangeRefall.Rows.Count, 1] as Excel.Range;
            rangeRef.EntireRow.Copy(misValue);
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", xlWorkSheetRef.Cells[rangeRefall.Rows.Count - 1, 1]) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            xlWorkSheetRef.SaveAs("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            releaseObject(xlAppRef);
            Thread.Sleep(3000);
            //////////////////////////////////////////////////////////////////////////////////
            xlApp2 = new Excel.ApplicationClass();
            xlApp2.Visible = true;
            xlApp2.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 = "[" + openfilex.Substring(7, 9) + "]";
            xlWorkBook = xlApp2.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            rCnt = range.Rows.Count;

            //CodeFinder cf;
            //cf = new CodeFinder(xlWorkBook, xlWorkSheet);
           // col3000 = cf.FindCodedColumn("3000", range);
           // col4000 = cf.FindCodedColumn("4000", range);
           // col5000 = cf.FindCodedColumn("5000", range);

            //col = cf.FindCodedColumn("8000", range);
            //col8000 = cf.FindCodedColumn("11000-1000", range);

            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "3000"))
                {
                    col3000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "4000"))
                {
                    col4000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "5000"))
                {
                    col5000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "8000"))//consigne de dégroupage
                {
                    col = cCnt;
                }
                if (Regex.Equals(valuecellabs, "11000-1000"))//consigne de suppresion
                {
                    col8000 = cCnt;
                    break;
                }

            }
            int row17000 = 0;
            //row17000 = cf.FindCodedRow("17000", range);


            cCnt = range.Columns.Count;
            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "17000"))
                {
                    row17000 = rCnt;
                    break;
                }
            }



            int fileflag = 0;
            for (int row = row17000+1; row <= values.GetUpperBound(0); row++)//20 pour annuel
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    xlWorkBookx1 = xlApp2.Workbooks.Open("D:\\ptw\\Histox.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Comptes annuels");
                    string[] namestable = { "ANNUEL-CR1.xlsx", "ANNUEL-CR2.xlsx", "ANNUEL-CR3.xlsx", "ANNUEL-BILACT1.xlsx", "ANNUEL-BILPAS1.xlsx", "ANNUEL-BILUSACT1.xlsx", "ANNUEL-BILUSPAS1.xlsx", "ANNUEL-FLUXFIN1.xlsx", "ANNUEL-FLUXTRES1.xlsx", "ANNUEL-RATIOS1.xlsx", "ANNUEL-RATIOS2.xlsx", "ANNUEL-SYNTH.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire
                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[row17000+1, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[row17000+1, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[row17000, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[row17000+1, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de row ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row17000 + 1, 1], xlWorkSheet.Cells[row - 1, col]) as Excel.Range;//A20
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());
                    Thread.Sleep(3000);
                    xlWorkBooknewx1 = xlApp2.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Comptes annuels");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);
                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp2.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Comptes annuels");
                    Excel.Range rangeremplace = xlWorkSheetremplace.UsedRange;
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas



                    //////////delete col8000 "-2"//////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    //for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                    //    if (Regex.Equals(valuedel, "-2"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                    //        rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    //        rangeremplace = xlWorkSheetremplace.UsedRange;
                    //        values8000 = (object[,])rangeremplace.Value2;
                    //        rowdel--;
                    //    }
                    //}
                    ///////////////row hide "-5"///////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }

                    ////////////////////////////////////////////////////////////////////////////
                    xlApp2.ActiveWindow.SplitRow = 0;
                    xlApp2.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    if (checkBox20.Checked == true)
                    {
                        fileAstyler = divisavenom;
                        Thread.Sleep(3000);
                        XmllireAnnuel(sender, e);
                    }

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = row17000+1;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp2.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp2);
        }

        private void diviserAnnuel(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            fichierprepare = textBox9.Text;// textBox9.Text;
            prefaceNP = "D:\\ptw\\Histo.xlsx";


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Range range = xlWorkSheet.UsedRange;
            string sss=range.Rows.Count.ToString();
            //Annu.Refer coller value
            Excel.Worksheet sheetAnnuRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Range rangeAnnuRefer = sheetAnnuRefer.UsedRange;
            rangeAnnuRefer.Copy(misValue);
            rangeAnnuRefer.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);



            //suppression des onglets
            

            Excel.Worksheet ComptesannuelRefssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Worksheet Comptesannuelssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");

           
            //Excel.Range rangeannuelA = Comptesannuelssheet.UsedRange.get_Range("A1", Comptesannuelssheet.Cells[Comptesannuelssheet.UsedRange.Rows.Count - 1, 1]) as Excel.Range;
            //Excel.Range rangeannuel1 = Comptesannuelssheet.UsedRange.get_Range("D1", Comptesannuelssheet.Cells[Comptesannuelssheet.UsedRange.Rows.Count - 1, 6]) as Excel.Range;
            //Excel.Range rangeannuel2 = Comptesannuelssheet.UsedRange.get_Range("K1", Comptesannuelssheet.Cells[Comptesannuelssheet.UsedRange.Rows.Count - 1, 13]) as Excel.Range;
            //Excel.Range rangeannuel3 = Comptesannuelssheet.UsedRange.get_Range("R1", Comptesannuelssheet.Cells[Comptesannuelssheet.UsedRange.Rows.Count - 1, 20]) as Excel.Range;

            //rangeannuelA.Copy(misValue);
            //rangeannuelA.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            //rangeannuel1.Copy(misValue);
            //rangeannuel1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            //rangeannuel2.Copy(misValue);
            //rangeannuel2.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            //rangeannuel3.Copy(misValue);
            //rangeannuel3.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


            Excel.Range rangeAnnuel1 = Comptesannuelssheet.UsedRange.get_Range("B1", "G1202") as Excel.Range;
            rangeAnnuel1.Copy(misValue);
            rangeAnnuel1.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Excel.Range rangeAnnuel2 = Comptesannuelssheet.UsedRange.get_Range("K1", "N1202") as Excel.Range;
            rangeAnnuel2.Copy(misValue);
            rangeAnnuel2.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Excel.Range rangeAnnuel3 = Comptesannuelssheet.UsedRange.get_Range("R1", "U1202") as Excel.Range;
            rangeAnnuel3.Copy(misValue);
            rangeAnnuel3.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Excel.Range rangeAnnuel4 = Comptesannuelssheet.UsedRange.get_Range("A1", "CC16") as Excel.Range;
            rangeAnnuel4.Copy(misValue);
            rangeAnnuel4.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            
            Excel.Worksheet delete2 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("Correctifs.Refer");
          //  Excel.Worksheet delete1 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("PreviNotaPme");
           // delete1.Delete();
            delete2.Delete();
            //Excel.Worksheet Historique = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            //Excel.Worksheet HistPrefac = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            //Excel.Worksheet HistCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            //Excel.Worksheet HistLangues = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues");
            //Excel.Worksheet HistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            //Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            //Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-s");
            //Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-s");
            //Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-s");
            //Excel.Worksheet HistRefersheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");
            //Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            //Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            //Excel.Worksheet Paramimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            //Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            //Excel.Worksheet Paramgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            //Excel.Worksheet AdminLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            //Excel.Worksheet AdminServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            //Excel.Worksheet Tsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            //Excel.Worksheet ParamSavsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            //Excel.Worksheet Macrossheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            //Excel.Worksheet Vsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            //Excel.Worksheet Mosaiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            //Excel.Worksheet GraphiquesSRsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            //Excel.Worksheet Graphimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            //Excel.Worksheet Dontdeletesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            //Excel.Worksheet Finsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            //Excel.Worksheet ChoixMethodessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            //Excel.Worksheet Noterecapitulativesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            //Excel.Worksheet SyntheseValorisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            //Excel.Worksheet DefinitionsArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            //Excel.Worksheet RappelRetraitementssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            //Excel.Worksheet RisqueEntreprisesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            //Excel.Worksheet ChoixTauxParamsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            //Excel.Worksheet TauxParamArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            //Excel.Worksheet CorrectifsSIGBilansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            //Excel.Worksheet APNNEsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            //Excel.Worksheet FiscaliteDiffereesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            //Excel.Worksheet PatrimonialAncAnccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            //Excel.Worksheet FondsDeCommercesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            //Excel.Worksheet Goodwillsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            //Excel.Worksheet AutresCapitalisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            //Excel.Worksheet Multiplessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            //Excel.Worksheet MethodesMixtessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            //Excel.Worksheet TransactionsComparablessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            //Excel.Worksheet GordonShapiroBatessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            //Excel.Worksheet CalculFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            //Excel.Worksheet DiscountedFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            //Excel.Worksheet CmpcWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            //Excel.Worksheet CmpcWaccArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            //Excel.Worksheet ModuleWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            //Excel.Worksheet CCEFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            //Excel.Worksheet TriRentabiliteProjetsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            //Excel.Worksheet TourDeTableSynthesesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            //Excel.Worksheet EvalLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            //Excel.Worksheet Controlessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            //Excel.Worksheet EvalServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            //Excel.Worksheet Composantessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            //Excel.Worksheet Jsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            //Excel.Worksheet Factgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            //Excel.Worksheet Lsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            //Excel.Worksheet Msheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            //Excel.Worksheet Tresoreriesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            //Excel.Worksheet ABsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            //Excel.Worksheet Paramtresorsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            //Excel.Worksheet Saisonnalitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            //Excel.Worksheet Zsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            //Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
            ////coller value de synthese

            //Historique.Delete();
            //HistPrefac.Delete();
            //HistCalculs.Delete();
            //HistLangues.Delete();
            //HistRefer.Delete();
            //Historiquesheet.Delete();
            //HistPrefacsheet.Delete();
            //HistCalculssheet.Delete();
            //HistLanguessheet.Delete();
            //HistRefersheet.Delete();
            //Identitesheet.Delete();
            //Paramimprsheet.Delete();
            //Psheet.Delete();
            //Paramgenerauxsheet.Delete();
            //AdminServicesheet.Delete();
            //Tsheet.Delete();
            //ParamSavsheet.Delete();
            //Macrossheet.Delete();
            //Vsheet.Delete();
            //Mosaiquesheet.Delete();
            //GraphiquesSRsheet.Delete();
            //Graphimprsheet.Delete();
            //Dontdeletesheet.Delete();
            //Finsheet.Delete();
            //ChoixMethodessheet.Delete();
            //Noterecapitulativesheet.Delete();
            //SyntheseValorisationssheet.Delete();
            //DefinitionsArrierePlansheet.Delete();
            //RappelRetraitementssheet.Delete();
            //RisqueEntreprisesheet.Delete();
            //ChoixTauxParamsheet.Delete();
            //TauxParamArrierePlansheet.Delete();
            //CorrectifsSIGBilansheet.Delete();
            //APNNEsheet.Delete();
            //FiscaliteDiffereesheet.Delete();
            //PatrimonialAncAnccsheet.Delete();
            //FondsDeCommercesheet.Delete();
            //Goodwillsheet.Delete();
            //AutresCapitalisationssheet.Delete();
            //Multiplessheet.Delete();
            //MethodesMixtessheet.Delete();
            //TransactionsComparablessheet.Delete();
            //GordonShapiroBatessheet.Delete();
            //CalculFCFsheet.Delete();
            //DiscountedFCFsheet.Delete();
            //CmpcWaccsheet.Delete();
            //CmpcWaccArrierePlansheet.Delete();
            //ModuleWaccsheet.Delete();
            //CCEFsheet.Delete();
            //TriRentabiliteProjetsheet.Delete();
            //TourDeTableSynthesesheet.Delete();
            //EvalLanguessheet.Delete();
            //Controlessheet.Delete();
            //EvalServicesheet.Delete();
            //Composantessheet.Delete();
            //Jsheet.Delete();
            //Factgenerauxsheet.Delete();
            //Lsheet.Delete();
            //Msheet.Delete();
            //Tresoreriesheet.Delete();
            //ABsheet.Delete();
            //Paramtresorsheet.Delete();
            //Saisonnalitesheet.Delete();
            //Zsheet.Delete();
            //model.Delete();
            Excel.Worksheet sheetCA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CA");
            Excel.Worksheet sheetInvestissements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investissements");
            Excel.Worksheet sheetCpteresultat = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Cpte Résultat");
            Excel.Worksheet sheetFinancements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Financements");
            Excel.Worksheet sheetbfr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("BFR");
            Excel.Worksheet sheetbilan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Bilan");
            Excel.Worksheet sheetcontrole2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles (2)");
            Excel.Worksheet sheetmultiple = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiple");
            Excel.Worksheet sheetvalo = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Valo et ouverture du capital");
            Excel.Worksheet sheetplan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Plan de financement");
            Excel.Worksheet sheetsynthese = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Synthèse SIG et SR");

            sheetCA.Delete();
            sheetInvestissements.Delete();
            sheetCpteresultat.Delete();
            sheetFinancements.Delete();
            sheetbfr.Delete();
            sheetbilan.Delete();
            sheetcontrole2.Delete();
            sheetmultiple.Delete();
            sheetvalo.Delete();
            sheetplan.Delete();
            sheetsynthese.Delete();
            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            supprimermoin2Annuel(sender, e);
            //subdiviser histo-s
            DiviStylerAnnuel(sender, e);

            int time2 = System.Environment.TickCount;
            int times = (time2 - time1) / 1000;
            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;
            timdiviserAnnuel = hours + " heures " + minuit + " minutes " + second;

            //timdiviserAnnuel = Convert.ToString(Convert.ToDecimal(times) / 1000);

            //MessageBox.Show("jobs done " + tim + " seconds used");
        }

        //private void XmllireAnnuel(object sender, EventArgs e)
        //{
        //    OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
        //    OpenFileDialog1.FileName = fileAstyler;
        //    OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
        //    OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
        //    //OpenFileDialog1.FilterIndex = 2;
        //    OpenFileDialog1.RestoreDirectory = true;
        //    if (OpenFileDialog1.FileName == "")
        //    {
        //        OpenFileDialog1.FileName = textBox14.Text.ToString();
        //        //OpenFileDialog1.ShowDialog();
        //    }
        //    try
        //    {
        //        textBox20.AppendText("==> Start Formatage des styles de COMPTES ANNUELS dans PrefaceNP : " + System.Environment.NewLine);
        //        Excel.Application xlapp = new Excel.ApplicationClass();
        //        xlapp.DisplayAlerts = false;
        //        xlapp.Application.DisplayAlerts = false;
        //        xlapp.Visible = true;

        //        int time1 = System.Environment.TickCount;

        //        Excel.Workbook xlworkbook = xlapp.Workbooks.Open(OpenFileDialog1.FileName.ToString());
        //        Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Comptes annuels");

        //        Excel.Workbook xlworkstyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");
        //        Excel.Worksheet xlstylesheet = (Excel.Worksheet)xlapp.Worksheets.get_Item("Annuel");

        //        Excel.Range rangeStyle = xlstylesheet.UsedRange;
        //        Excel.Range rangeToChange = xlworksheet.UsedRange;
        //        rangeToChange.ClearFormats();
        //        object[,] ValuesStyle = (object[,])rangeStyle.Value2;
        //        object[,] values = (object[,])rangeToChange.Value2;

        //        int col27000style = 0;
        //        int col90001000 = 0;

        //        for (int i = 1; i <= rangeStyle.Columns.Count; i++)
        //        {
        //            if (ValuesStyle[rangeStyle.Rows.Count, i] != null)
        //            {
        //                if (ValuesStyle[rangeStyle.Rows.Count, i].ToString() == "27000")
        //                {
        //                    col27000style = i;
        //                    break;
        //                }
        //            }
        //        }

        //        for (int i = 1; i <= rangeToChange.Columns.Count; i++)
        //        {
        //            if (values[rangeToChange.Rows.Count, i] != null)
        //            {
        //                if (values[rangeToChange.Rows.Count, i].ToString() == "9000-1000")
        //                {
        //                    col90001000 = i;
        //                    break;
        //                }
        //            }
        //        }
        //        for (int i = 1; i <= 20; i++)
        //        {
        //            if (values[i, col90001000] == null)
        //            {
        //                Excel.Range rangeTohide = xlworksheet.Cells[i, 1] as Excel.Range;
        //                rangeTohide.EntireRow.Hidden = true;
        //            }
        //        }
        //        rangeToChange.ClearFormats();
        //        for (int i = 7; i <= 43; i++)
        //        {
        //            Excel.Range rangeCherche = rangeStyle.get_Range("A" + i, "A" + i);

        //            if (rangeCherche.Value2 != null)
        //            {
        //                string cherche = rangeCherche.Value2.ToString();
        //                string fontSize = "0";
        //                for (int t = 1; t <= rangeToChange.Rows.Count; t++)
        //                {
        //                    if (values[t, col90001000] != null)
        //                    {
        //                        if (cherche == values[t, col90001000].ToString())
        //                        {
        //                            Excel.Range rangeCopy = rangeStyle.get_Range("C" + i, "Z" + i);
        //                            Excel.Range range1 = rangeToChange.get_Range("A" + t, "AD" + t);
        //                            Excel.Range rangeFontSize = rangeStyle.get_Range(rangeStyle.Cells[i, col27000style], rangeStyle.Cells[i, col27000style]);
        //                            fontSize = rangeFontSize.Value2.ToString();

        //                            rangeCopy.Copy();
        //                            range1.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        //                            try
        //                            {
        //                                range1.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        //                            }
        //                            catch (Exception exf)
        //                            {
        //                            }
        //                            range1.Font.Size = int.Parse(fontSize);
        //                        }
        //                    }
        //                }
        //            }

        //        }
        //        Excel.Range range = xlworksheet.UsedRange;
        //        object[,] values1 = (object[,])range.Value2;

        //        int col12000 = 0;
        //        int row763000 = 0;
        //        for (int i = 1; i <= range.Columns.Count; i++)
        //        {
        //            if (values1[range.Rows.Count, i] != null)
        //            {
        //                if (values1[range.Rows.Count, i].ToString() == "12000")
        //                {
        //                    col12000 = i;
        //                    break;

        //                }
        //            }
        //        }
        //        for (int i = 1; i <= range.Rows.Count; i++)
        //        {
        //            if (values1[i, range.Columns.Count] != null)
        //            {
        //                if (values1[i, range.Columns.Count].ToString() == "763000")
        //                {
        //                    row763000 = i;
        //                    break;

        //                }
        //            }
        //        }


        //        for (int i = 1; i <= 20; i++)
        //        {
        //            if (values1[i, col90001000] == null)
        //            {
        //                Excel.Range rangeTohide = xlworksheet.Cells[i, 1] as Excel.Range;
        //                rangeTohide.EntireRow.Hidden = true;
        //            }
        //        }

        //        for (int i = 1; i <= row763000; i++)
        //        {
        //            if (values1[i, col12000] != null && values1[i, col12000].ToString() != "1")
        //            {
        //                Excel.Range rangeTohide = xlworksheet.Cells[i, 1] as Excel.Range;
        //                //rangeTohide.EntireRow.Hidden = true;
        //            }
        //        }
        //        xlworkbook.Save();
        //        xlapp.Quit();




        //    }
        //    catch (Exception ex)
        //    {
        //        textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
        //        // }
        //        // ////////////////open excel////////////////////////
        //        // Excel.Application xlApp;
        //        // Excel.Workbook xlWorkBook;
        //        // object misValue = System.Reflection.Missing.Value;
        //        // xlApp = new Excel.ApplicationClass();
        //        // xlApp.Visible = true;
        //        // xlApp.DisplayAlerts = false;


        //        // string openfilex = OpenFileDialog1.FileName.ToString();
        //        // xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //        // //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
        //        // //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
        //        // Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Comptes annuels"];
        //        // Excel.Range range = xlWorkSheet.UsedRange;
        //        // object[,] values = (object[,])range.Value2;

        //        // //////////////////////////////open le fichier style XML//////////////////////
        //        // //OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
        //        // //OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
        //        // //OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
        //        // //OpenFileDialog2.ShowDialog();
        //        // //stylexml = OpenFileDialog2.FileName.ToString();

        //        // if (textBox10.Text != null)
        //        //     stylexml = textBox10.Text;
        //        // else
        //        //     MessageBox.Show("Veuillez choiser le fichier style en format XML");


        //        // XmlDocument appstyleDoc = new XmlDocument();
        //        // appstyleDoc.Load(stylexml);
        //        // //appstyleDoc.Load("D:\\appstyle22.xml");

        //        // /////////////////////////////////////set palette couleur///////////////////////////
        //        // xlWorkBook.ResetColors();
        //        // XmlElement indexxmlelement = appstyleDoc.DocumentElement;
        //        // XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
        //        // XmlNode indexstylenode = indexstylenodelist.Item(0);

        //        // //for (int nindex = 1; nindex <= 56; nindex++)
        //        // //{
        //        // //    string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
        //        // //    int valeur2index = Convert.ToInt32(valeurindex);
        //        // //    xlWorkBook.set_Colors(nindex, valeur2index);
        //        // //}
        //        // range.EntireRow.Font.Size = 8;
        //        // //range.Rows.AutoFit();
        //        // //Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //        // //rangemasquer.EntireRow.Hidden = true;
        //        // ////////////////////////////////////////////////////////////////////////////////////

        //        // int rCnt = 0;
        //        // int cCnt = 0;
        //        // //int col = 0;
        //        // int col15000 = 0;
        //        // int col11000 = 0;
        //        // rCnt = range.Rows.Count;

        //        //// CodeFinder cf;
        //        //// cf = new CodeFinder(xlWorkBook, xlWorkSheet);
        //        //// col15000 = cf.FindCodedColumn("9000-1000", range);


        //        // for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
        //        // {
        //        //     string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
        //        //     if (Regex.Equals(valuecellabs, "9000-1000"))
        //        //     {
        //        //         col15000 = cCnt;

        //        //     }
        //        //     if (Regex.Equals(valuecellabs, "11000-1000"))
        //        //     {
        //        //         col11000 = cCnt;
        //        //         break;
        //        //     }
        //        // }



        //        // ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
        //        // XmlElement nbstyle = appstyleDoc.DocumentElement;
        //        // XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
        //        // XmlNode nbstylenode = nbstylelist.Item(0);

        //        // string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
        //        // int nbtotalint = 120;
        //        // nbtotalint = Convert.ToInt32(nbtotal);
        //        // string[] tablestyle = new string[nbtotalint+1];
        //        // for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //        // {
        //        //     tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
        //        // }
        //        // /////////////////////////////////////////////////////////////////////////////////////////////


        //        // int row = 1;
        //        // string colcount = "";
        //        // int time1 = System.Environment.TickCount;
        //        // int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
        //        // for (row = 1; row <= rowCountx - 1; row++)
        //        // {
        //        //     string value = Convert.ToString(values[row, col15000]);
        //        //     for (int nbs = 1; nbs <= nbtotalint; nbs++)
        //        //     {
        //        //         if (Regex.Equals(value, tablestyle[nbs]))
        //        //         {
        //        //             XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
        //        //             if (xstyle != null)
        //        //             {
        //        //                 colcount = (xstyle.SelectSingleNode("col")).InnerText;
        //        //             }
        //        //             int colcountx = Convert.ToInt32(colcount);
        //        //             for (int colc = 1; colc <= colcountx; colc++)
        //        //             {
        //        //                 XmlElement xmlelement = appstyleDoc.DocumentElement;
        //        //                 XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
        //        //                 XmlNode stylenode = stylenodelist.Item(0);
        //        //                 string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
        //        //                 string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
        //        //                 //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
        //        //                 //int colorBx = Convert.ToInt32(colorB);
        //        //                 //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
        //        //                 string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
        //        //                // string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

        //        //                 string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
        //        //                 string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
        //        //                 string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

        //        //                 string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
        //        //                 string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
        //        //                 string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
        //        //                 string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
        //        //                 string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
        //        //                 string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
        //        //                 string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
        //        //                 string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
        //        //                 string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
        //        //                 string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

        //        //                 string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
        //        //                 string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
        //        //                 string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
        //        //                 string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
        //        //                 string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //        //                 int intmergecellcount = Convert.ToInt32(mergecellcount);

        //        //                 string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
        //        //                 string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
        //        //                 string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
        //        //                 string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
        //        //                 string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
        //        //                 ///////////////////////////////////merge process///////////////////////////////////////////
        //        //                 if (mergecell == "True")
        //        //                 {
        //        //                     if (intmergecellcount > 1)
        //        //                     {
        //        //                         Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
        //        //                         rangemerge.Merge(false);
        //        //                         //rangemerge.HorizontalAlignment = 1;

        //        //                         for (int countarea = 1; countarea < intmergecellcount; countarea++)
        //        //                         {
        //        //                             XmlElement mergexmlelement = appstyleDoc.DocumentElement;
        //        //                             int mergecolindex = colc + countarea;
        //        //                             XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
        //        //                             XmlNode mergestylenode = mergestylenodelist.Item(0);
        //        //                             mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
        //        //                             appstyleDoc.Save(stylexml);
        //        //                         }
        //        //                     }
        //        //                 }
        //        //                 /////////////////////////////exception traitement/////////////////////////////////
        //        //                 //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
        //        //                 //xlWorkSheet.Cells.ColumnWidth = 20;
        //        //                 //////////////////////////////////////////////////////////////////////////////////

        //        //                 /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
        //        //                 Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
        //        //                 rangeDelx.Font.Name = fontname;
        //        //                 rangeDelx.Font.Size = Convert.ToInt32(fontsize);
        //        //                 //2003-2010
        //        //                 rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
        //        //                 //rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
        //        //                 //rangeDelx.Value2 = fontcolorindex;

        //        //                 rangeDelx.Font.Bold = (fontbold == "True");
        //        //                 rangeDelx.Font.Italic = (fontitalic == "True");
        //        //                 rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
        //        //                 //rangeDelx.Value2 += "bgcolor" + bgcolorindex;
        //        //                 rangeDelx.Interior.Color = Convert.ToInt32(bgcolor);
        //        //                 //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
        //        //                 rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

        //        //                 rangeDelx.WrapText = (wraptext == "True");
        //        //                 rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
        //        //                 rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

        //        //                 /////////////////////////////////////////////////////////////////////////////////////////
        //        //                 mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
        //        //                 //ne peut pas modifier les cellules fusionner
        //        //                 if (mergecellcount == "False")
        //        //                 {
        //        //                     rangeDelx.NumberFormat = nomberformat;
        //        //                     rangeDelx.Locked = (locked == "True");
        //        //                     rangeDelx.Locked = (formulahidden == "True");
        //        //                 }
        //        //                 ///////////////////////////////////////////////////////////////////////////////////////////
        //        //                 rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
        //        //                 rangeDelx.RowHeight = Convert.ToDouble(rowheight);
        //        //             }
        //        //         }
        //        //     }
        //        // }
        //        // xlApp.ActiveWindow.DisplayGridlines = false;
        //        // //range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        //        // //range.Rows.AutoFit();
        //        // //Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
        //        // //rangemasquer2.EntireRow.Hidden = true;

        //        // //pour consigne de masquage
        //        // //Excel.Range rangeremplace = xlWorkSheet.UsedRange;
        //        // //object[,] values8000 = (object[,])rangeremplace.Value2;
        //        // //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
        //        // //{
        //        // //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
        //        // //    if (Regex.Equals(valuedel, "-1"))
        //        // //    {
        //        // //        Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
        //        // //        rangeDely.EntireRow.Hidden = true;
        //        // //    }
        //        // //}

        //        // Excel.Range rangeremplace = xlWorkSheet.UsedRange;
        //        // object[,] values8000 = (object[,])rangeremplace.Value2;
        //        // ///////////////row hide "-5"////////////////////////////////////////////////
        //        // for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
        //        // {
        //        //     string valuedel = Convert.ToString(values8000[rowhide, col11000]);
        //        //     if (Regex.Equals(valuedel, "-5"))
        //        //     {
        //        //         Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col11000] as Excel.Range;
        //        //         rangeDely.EntireRow.Hidden = true;
        //        //     }
        //        // }

        //        // Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("Y1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
        //        // Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
        //        //// Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;

        //        // //Excel.Range rangeHideB = xlWorkSheet.UsedRange.Columns[2] as Excel.Range;
        //        // //Excel.Range rangeHideC = xlWorkSheet.UsedRange.Columns[3] as Excel.Range;
        //        // //rangeHideB.Hidden = true;
        //        // //rangeHideC.Hidden = true;
        //        // //consigne supression
        //        // //rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
        //        // //rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //        // //rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        //        // rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
        //        // rangeDelete.EntireColumn.Hidden = true;


        //        // int time2 = System.Environment.TickCount;
        //        // int times = time2 - time1;
        //        // string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
        //        // //MessageBox.Show("jobs done " + tim + " seconds used");
        //        // //xlWorkBook.Save();
        //        // if (pathstylerfinal != null) System.IO.Directory.CreateDirectory(pathstylerfinal);
        //        // if (divitylerfinal != null) xlWorkBook.SaveCopyAs(divitylerfinal);
        //        // if (divitylerfinal != null) xlWorkBook.Close(false, misValue, misValue);

        //        // if (divitylerfinal == null) xlWorkBook.Close(true, misValue, misValue);
        //        // xlApp.Quit();

        //        // releaseObject(xlWorkSheet);
        //        // releaseObject(xlWorkBook);
        //        // releaseObject(xlApp);
        //    }
        //}
        private void XmllireAnnuel(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.FileName = fileAstyler;
            OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
            OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            //OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;
            if (OpenFileDialog1.FileName == "")
            {
                OpenFileDialog1.FileName = textBox14.Text.ToString();
                //OpenFileDialog1.ShowDialog();
            }

            ////////////////open excel////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;


            string openfilex = OpenFileDialog1.FileName.ToString();
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Comptes annuels"];
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////
            //OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
            //OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
            //OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
            //OpenFileDialog2.ShowDialog();
            //stylexml = OpenFileDialog2.FileName.ToString();

            if (textBox10.Text != null)
                stylexml = textBox10.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");


            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);
            //appstyleDoc.Load("D:\\appstyle22.xml");

            /////////////////////////////////////set palette couleur///////////////////////////
            xlWorkBook.ResetColors();
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);

            //for (int nindex = 1; nindex <= 56; nindex++)
            //{
            //    string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
            //    int valeur2index = Convert.ToInt32(valeurindex);
            //    xlWorkBook.set_Colors(nindex, valeur2index);
            //}
            range.EntireRow.Font.Size = 8;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer.EntireRow.Hidden = true;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;
            int col15000 = 0;
            int col11000 = 0;
            rCnt = range.Rows.Count;

            // CodeFinder cf;
            // cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            // col15000 = cf.FindCodedColumn("9000-1000", range);


            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "9000-1000"))
                {
                    col15000 = cCnt;

                }
                if (Regex.Equals(valuecellabs, "11000-1000"))
                {
                    col11000 = cCnt;
                    break;
                }
            }



            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint + 1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////


            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;
            int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
            for (row = 1; row <= rowCountx - 1; row++)
            {
                string value = Convert.ToString(values[row, col15000]);
                for (int nbs = 1; nbs <= nbtotalint; nbs++)
                {
                    if (Regex.Equals(value, tablestyle[nbs]))
                    {
                        XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                        if (xstyle != null)
                        {
                            colcount = (xstyle.SelectSingleNode("col")).InnerText;
                        }
                        int colcountx = Convert.ToInt32(colcount);
                        for (int colc = 1; colc <= colcountx; colc++)
                        {
                            XmlElement xmlelement = appstyleDoc.DocumentElement;
                            XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                            XmlNode stylenode = stylenodelist.Item(0);
                            string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                            string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                            //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                            //int colorBx = Convert.ToInt32(colorB);
                            //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                            string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                            // string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                            string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                            string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                            string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                            string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
                            string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                            string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                            string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                            string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                            string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                            string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                            string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                            string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                            string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                            string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                            string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                            string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                            string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                            string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            int intmergecellcount = Convert.ToInt32(mergecellcount);

                            string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                            string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                            string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                            string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                            string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
                            ///////////////////////////////////merge process///////////////////////////////////////////
                            if (mergecell == "True")
                            {
                                if (intmergecellcount > 1)
                                {
                                    Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
                                    rangemerge.Merge(false);
                                    //rangemerge.HorizontalAlignment = 1;

                                    for (int countarea = 1; countarea < intmergecellcount; countarea++)
                                    {
                                        XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                        int mergecolindex = colc + countarea;
                                        XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                        XmlNode mergestylenode = mergestylenodelist.Item(0);
                                        mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                        appstyleDoc.Save(stylexml);
                                    }
                                }
                            }
                            /////////////////////////////exception traitement/////////////////////////////////
                            //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                            //xlWorkSheet.Cells.ColumnWidth = 20;
                            //////////////////////////////////////////////////////////////////////////////////

                            /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
                            Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                            rangeDelx.Font.Name = fontname;
                            rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                            //2003-2010
                            rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
                            //rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                            //rangeDelx.Value2 = fontcolorindex;

                            rangeDelx.Font.Bold = (fontbold == "True");
                            rangeDelx.Font.Italic = (fontitalic == "True");
                            rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                            //rangeDelx.Value2 += "bgcolor" + bgcolorindex;
                            rangeDelx.Interior.Color = Convert.ToInt32(bgcolor);
                            //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                            rangeDelx.WrapText = (wraptext == "True");
                            rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                            rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

                            /////////////////////////////////////////////////////////////////////////////////////////
                            mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            //ne peut pas modifier les cellules fusionner
                            if (mergecellcount == "1")
                            {
                                try
                                {
                                    rangeDelx.NumberFormat = nomberformat;
                                    try
                                    {
                                        rangeDelx.Locked = (locked == "True");
                                        rangeDelx.Locked = (formulahidden == "True");
                                    }
                                    catch
                                    {
                                    }
                                }
                                catch { 
                                }
                             }
                            ///////////////////////////////////////////////////////////////////////////////////////////
                            rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                            rangeDelx.RowHeight = Convert.ToDouble(rowheight);
                        }
                    }
                }
            }
            xlApp.ActiveWindow.DisplayGridlines = false;
            //range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer2.EntireRow.Hidden = true;

            //pour consigne de masquage
            //Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            //object[,] values8000 = (object[,])rangeremplace.Value2;
            //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            //{
            //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
            //    if (Regex.Equals(valuedel, "-1"))
            //    {
            //        Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
            //        rangeDely.EntireRow.Hidden = true;
            //    }
            //}

            Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            object[,] values8000 = (object[,])rangeremplace.Value2;
            ///////////////row hide "-5"////////////////////////////////////////////////
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col11000]);
                if (Regex.Equals(valuedel, "-5"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col11000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }

            Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("Y1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
            Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
            // Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;

            //Excel.Range rangeHideB = xlWorkSheet.UsedRange.Columns[2] as Excel.Range;
            //Excel.Range rangeHideC = xlWorkSheet.UsedRange.Columns[3] as Excel.Range;
            //rangeHideB.Hidden = true;
            //rangeHideC.Hidden = true;
            //consigne supression
            //rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            //rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
            rangeDelete.EntireColumn.Hidden = true;


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            //xlWorkBook.Save();
            if (pathstylerfinal != null) System.IO.Directory.CreateDirectory(pathstylerfinal);
            if (divitylerfinal != null) xlWorkBook.SaveCopyAs(divitylerfinal);
            if (divitylerfinal != null) xlWorkBook.Close(false, misValue, misValue);

            if (divitylerfinal == null) xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        #endregion

        #region Synthese
        //////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////////////////Eval.ptw///////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////


        //Synthese     SynthèseValorisations
        private void supprimermoin2Synthese(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////116000//////////////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row1012000 = 0;

            cCnt = range.Columns.Count;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            row1012000 = cf.FindCodedRow("116000", range);


            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "116000"))
            //    {
            //        row1012000 = rCnt;
            //        break;
            //    }
            //}

            for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            {
                string value = Convert.ToString(values[row1012000, col]);
                if (Regex.Equals(value, "-2"))
                {
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row1012000, col] as Excel.Range;
                    rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    col--;
                }
            }

            range = xlWorkSheet.UsedRange;
            cCnt = range.Columns.Count;
            values = (object[,])range.Value2;
            for (int col = 1; col <= cCnt; col++)
            {
                string value = Convert.ToString(values[row1012000, col]);
                if (Regex.Equals(value, "-4"))
                {
                    Excel.Range rangeEffacer = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col], xlWorkSheet.Cells[row1012000 - 1, col]) as Excel.Range;
                    rangeEffacer.ClearContents();
                }
            }



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void diviserSynthese(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            fichierprepare = textBox9.Text;// textBox9.Text;
            prefaceNP = "D:\\ptw\\Histo.xlsx";


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Range range = xlWorkSheet.UsedRange;



            //suppression des onglets
            Excel.Worksheet Historique = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet HistPrefac = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet HistCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Worksheet HistLangues = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues");
            Excel.Worksheet HistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");


            Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-s");
            Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-s");
            Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-s");
            Excel.Worksheet HistRefersheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");


            Excel.Worksheet ComptesannuelRefssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Worksheet Comptesannuelssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            Excel.Worksheet Paramimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            Excel.Worksheet Paramgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            Excel.Worksheet AdminLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            Excel.Worksheet AdminServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            Excel.Worksheet Tsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            Excel.Worksheet ParamSavsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            Excel.Worksheet Macrossheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            Excel.Worksheet Vsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            Excel.Worksheet Mosaiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            Excel.Worksheet GraphiquesSRsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            Excel.Worksheet Graphimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            Excel.Worksheet Dontdeletesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            Excel.Worksheet Finsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            Excel.Worksheet ChoixMethodessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            Excel.Worksheet Noterecapitulativesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            Excel.Worksheet SyntheseValorisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Worksheet DefinitionsArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            Excel.Worksheet RappelRetraitementssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            Excel.Worksheet RisqueEntreprisesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            Excel.Worksheet ChoixTauxParamsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            Excel.Worksheet TauxParamArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            Excel.Worksheet CorrectifsSIGBilansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            Excel.Worksheet APNNEsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            Excel.Worksheet FiscaliteDiffereesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            Excel.Worksheet PatrimonialAncAnccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            Excel.Worksheet FondsDeCommercesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            Excel.Worksheet Goodwillsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            Excel.Worksheet AutresCapitalisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            Excel.Worksheet Multiplessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            Excel.Worksheet MethodesMixtessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            Excel.Worksheet TransactionsComparablessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            Excel.Worksheet GordonShapiroBatessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            Excel.Worksheet CalculFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            Excel.Worksheet DiscountedFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            Excel.Worksheet CmpcWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            Excel.Worksheet CmpcWaccArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            Excel.Worksheet ModuleWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            Excel.Worksheet CCEFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            Excel.Worksheet TriRentabiliteProjetsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            Excel.Worksheet TourDeTableSynthesesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            Excel.Worksheet EvalLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            Excel.Worksheet Controlessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            Excel.Worksheet EvalServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            Excel.Worksheet Composantessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            Excel.Worksheet Jsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            Excel.Worksheet Factgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            Excel.Worksheet Lsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            Excel.Worksheet Msheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            Excel.Worksheet Tresoreriesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            Excel.Worksheet ABsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            Excel.Worksheet Paramtresorsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            Excel.Worksheet Saisonnalitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            Excel.Worksheet Zsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
            //coller value de synthese
            Excel.Range rangesynthese = SyntheseValorisationssheet.UsedRange;
            Excel.Worksheet deleteworksheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Histo.Macros-s");
            Excel.Worksheet deleteworksheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Typologies IFRS-s");
            rangesynthese.Copy(misValue);
            rangesynthese.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            Excel.Worksheet sheetCA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CA");
            Excel.Worksheet sheetInvestissements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investissements");
            Excel.Worksheet sheetCpteresultat = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Cpte Résultat");
            Excel.Worksheet sheetFinancements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Financements");
            Excel.Worksheet sheetbfr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("BFR");
            Excel.Worksheet sheetbilan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Bilan");
            Excel.Worksheet sheetcontrole2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles (2)");
            Excel.Worksheet sheetmultiple = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiple");
            Excel.Worksheet sheetvalo = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Valo et ouverture du capital");
            Excel.Worksheet sheetplan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Plan de financement");
            Excel.Worksheet sheetsynthese = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Synthèse SIG et SR");

            sheetCA.Delete();
            sheetInvestissements.Delete();
            sheetCpteresultat.Delete();
            sheetFinancements.Delete();
            sheetbfr.Delete();
            sheetbilan.Delete();
            sheetcontrole2.Delete();
            sheetmultiple.Delete();
            sheetvalo.Delete();
            sheetplan.Delete();
            sheetsynthese.Delete();
           // Excel.Worksheet delete1 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("PreviNotaPme");
          //  delete1.Delete();
            deleteworksheet1.Delete();
            deleteworksheet2.Delete();
            model.Delete();
            Historique.Delete();
            HistPrefac.Delete();
            HistCalculs.Delete();
            HistLangues.Delete();
            HistRefer.Delete();

            Historiquesheet.Delete();
            HistPrefacsheet.Delete();
            HistCalculssheet.Delete();
            HistLanguessheet.Delete();
            HistRefersheet.Delete();

            ComptesannuelRefssheet.Delete();
            Comptesannuelssheet.Delete();
            Osheet.Delete();
            Identitesheet.Delete();
            Paramimprsheet.Delete();
            Psheet.Delete();
            Paramgenerauxsheet.Delete();
            AdminLanguessheet.Delete();
            AdminServicesheet.Delete();
            Tsheet.Delete();
            ParamSavsheet.Delete();
            Macrossheet.Delete();
            Vsheet.Delete();
            Mosaiquesheet.Delete();
            GraphiquesSRsheet.Delete();
            Graphimprsheet.Delete();
            Dontdeletesheet.Delete();
            Finsheet.Delete();
            ChoixMethodessheet.Delete();
            Noterecapitulativesheet.Delete();
            //SyntheseValorisationssheet.Delete();
            DefinitionsArrierePlansheet.Delete();
            RappelRetraitementssheet.Delete();
            RisqueEntreprisesheet.Delete();
            ChoixTauxParamsheet.Delete();
            TauxParamArrierePlansheet.Delete();
            CorrectifsSIGBilansheet.Delete();
            APNNEsheet.Delete();
            FiscaliteDiffereesheet.Delete();
            PatrimonialAncAnccsheet.Delete();
            FondsDeCommercesheet.Delete();
            Goodwillsheet.Delete();
            AutresCapitalisationssheet.Delete();
            Multiplessheet.Delete();
            MethodesMixtessheet.Delete();
            TransactionsComparablessheet.Delete();
            GordonShapiroBatessheet.Delete();
            CalculFCFsheet.Delete();
            DiscountedFCFsheet.Delete();
            CmpcWaccsheet.Delete();
            CmpcWaccArrierePlansheet.Delete();
            ModuleWaccsheet.Delete();
            CCEFsheet.Delete();
            TriRentabiliteProjetsheet.Delete();
            TourDeTableSynthesesheet.Delete();
            EvalLanguessheet.Delete();
            Controlessheet.Delete();
            EvalServicesheet.Delete();
            Composantessheet.Delete();
            Jsheet.Delete();
            Factgenerauxsheet.Delete();
            Lsheet.Delete();
            Msheet.Delete();
            Tresoreriesheet.Delete();
            ABsheet.Delete();
            Paramtresorsheet.Delete();
            Saisonnalitesheet.Delete();
            Zsheet.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            supprimermoin2Synthese(sender, e);
            //subdiviser histo-s
            DiviStylerSynthese(sender, e);

            int time2 = System.Environment.TickCount;

            int times = (time2 - time1) / 1000;
            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;
            timdiviserSynthese = hours + " heures " + minuit + " minutes " + second;

            //timdiviserSynthese = Convert.ToString(Convert.ToDecimal(times) / 1000);

            //MessageBox.Show("jobs done " + tim + " seconds used");
        }

        private void DiviStylerSynthese(object sender, EventArgs e)
        {
            pathnotapme = textBox3.Text;
            pathstylerfinal = textBox6.Text;

            string openfilex = "D:\\ptw\\Histo.xlsx";

            ////////////////open excel///////////////////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("SynthèseValorisations");
            Excel.Range rangeRefall = xlWorkSheetRef.UsedRange;
            //exception!!!
            xlWorkSheetRef.Cells.ColumnWidth = 20;

            Excel.Range rangeRef = xlWorkSheetRef.Cells[rangeRefall.Rows.Count, 1] as Excel.Range;
            rangeRef.EntireRow.Copy(misValue);
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", xlWorkSheetRef.Cells[rangeRefall.Rows.Count - 1, 1]) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            xlWorkSheetRef.SaveAs("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            //////////////////////////////////////////////////////////////////////////////////
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 = "[" + openfilex.Substring(7, 9) + "]";
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            rCnt = range.Rows.Count;


            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col3000 = cf.FindCodedColumn("3000", range);
            col4000 = cf.FindCodedColumn("4000", range);
            col5000 = cf.FindCodedColumn("5000", range);
            col8000 = cf.FindCodedColumn("41000", range);
            col = cf.FindCodedColumn("42000", range);

            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "3000"))
            //    {
            //        col3000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "4000"))
            //    {
            //        col4000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "5000"))
            //    {
            //        col5000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "41000"))//consogne de suppresion
            //    {
            //        col8000 = cCnt;
            //    }
            //    if (Regex.Equals(valuecellabs, "42000"))//consigne de dégroupage
            //    {
            //        col = cCnt;
            //        break;
            //    }
            //}



            int row5000 = 0;
            cCnt = range.Columns.Count;

            row5000 = cf.FindCodedRow("4400-6000", range);//assurer que ce ligne est toujour avant la primiere bloc
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "4400-5000"))
            //    {
            //        row5000 = rCnt;
            //        break;
            //    }
            //}






            int fileflag = 0;
            for (int row = row5000+1; row <= values.GetUpperBound(0); row++)//19 pour annuel //8 pour synthese
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("SynthèseValorisations");
                    string[] namestable = { "EVAL-SYNTHVALO2.xlsx", "EVAL-SYNTHVALO1.xlsx", "EVAL-SYNTHMULT1.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire
                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[row5000+1, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[row5000+1, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[row5000, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[row5000+1, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de row ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row5000 + 1, 1], xlWorkSheet.Cells[row - 1, col]) as Excel.Range;//A21
                    
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("SynthèseValorisations");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);
                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    //xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                   // Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("SynthèseValorisations");
                    Excel.Range rangeremplace = xlWorkSheetremplace.UsedRange;
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas



                    //////////delete col8000 "-2"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    //for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                    //    if (Regex.Equals(valuedel, "-2"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                    //        rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    //        rangeremplace = xlWorkSheetremplace.UsedRange;
                    //        values8000 = (object[,])rangeremplace.Value2;
                    //        rowdel--;
                    //    }
                    //}
                    /////////////////row hide "-5"////////////////////////////////////////////////
                    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                    //    if (Regex.Equals(valuedel, "-5"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                    //        rangeDely.EntireRow.Hidden = true;
                    //    }
                    //}
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }

                    //object[,] valuesNX = (object[,])rangeremplace.Value2;
                    ////string valueNX = Convert.ToString(valuesNX[row, col]);
                    //for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row3000, col8000]) != "-7")//-7 non zero
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}
                    //for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row4000, col8000]) != "-7")//-7 non zero
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}
                    //for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    //{
                    //    Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                    //    {
                    //        rangeprey.Locked = false;
                    //        rangeprey.FormulaHidden = false;
                    //    }
                    //    if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                    //    {
                    //        rangeprey.Value2 = 0;
                    //        rangeprey.Locked = true;
                    //        rangeprey.FormulaHidden = true;
                    //    }
                    //    Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                    //    if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row5000, col8000]) != "-7")//-7 non zero
                    //    {
                    //        rangeDely.Value2 = 0;
                    //    }
                    //}

                    ////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    if (checkBox20.Checked == true)
                    {
                        fileAstyler = divisavenom;
                        XmllireSynthese(sender, e);
                    }

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = row5000 + 1;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void XmllireSynthese(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.FileName = fileAstyler;
            OpenFileDialog1.InitialDirectory = "D:\\ptw\\";
            OpenFileDialog1.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            //OpenFileDialog1.FilterIndex = 2;
            OpenFileDialog1.RestoreDirectory = true;
            if (OpenFileDialog1.FileName == "")
            {
                OpenFileDialog1.FileName = textBox14.Text.ToString();
                //OpenFileDialog1.ShowDialog();
            }

            ////////////////open excel////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;


            string openfilex = OpenFileDialog1.FileName.ToString();
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
           // xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////
            //OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
            //OpenFileDialog2.InitialDirectory = "D:\\ptw\\";
            //OpenFileDialog2.Filter = "XML fichier .xml|*.xml";
            //OpenFileDialog2.ShowDialog();
            //stylexml = OpenFileDialog2.FileName.ToString();

            if (textBox10.Text != null)
                stylexml = textBox10.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");


            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);
            //appstyleDoc.Load("D:\\appstyle22.xml");

            /////////////////////////////////////set palette couleur///////////////////////////
            xlWorkBook.ResetColors();
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);

            //for (int nindex = 1; nindex <= 56; nindex++)
            //{
            //    string valeurindex = indexstylenode.SelectNodes("index" + nindex).Item(0).InnerText.ToString();
            //    int valeur2index = Convert.ToInt32(valeurindex);
            //    xlWorkBook.set_Colors(nindex, valeur2index);
            //}
            range.EntireRow.Font.Size = 3;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer.EntireRow.Hidden = true;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;
            int col15000 = 0;

            rCnt = range.Rows.Count;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col15000 = cf.FindCodedColumn("40000", range);


            //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "40000")) //40000 pour synthese
            //    {
            //        col15000 = cCnt;
            //        break;
            //    }
            //}



            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint+1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////

            //range.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlDiagonalDown].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlDiagonalUp].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlHairline;

            //range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlHairline;

            //range.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlHairline, Excel.XlColorIndex.xlColorIndexAutomatic, misValue);
            /////////////////////////////////////////////////////////////////////////////////////////////



            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;
            int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
            for (row = 1; row <= rowCountx - 1; row++)
            {
                Excel.Range rangeline = xlWorkSheet.UsedRange.Rows[row] as Excel.Range;

                rangeline.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeline.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlHairline;

                rangeline.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeline.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlHairline;

                rangeline.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeline.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlHairline;

                rangeline.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                rangeline.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlHairline;
                
                    string value = Convert.ToString(values[row, col15000]);
                    for (int nbs = 1; nbs <= nbtotalint; nbs++)
                    {
                        if (Regex.Equals(value, tablestyle[nbs]))
                        {
                            XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                            if (xstyle != null)
                            {
                                colcount = (xstyle.SelectSingleNode("col")).InnerText;
                            }
                            int colcountx = Convert.ToInt32(colcount);
                            for (int colc = colcountx; colc >= 1; colc--)//for (int colc = colcountx; colc >= 1; colc--)//for (int colc = 1; colc <= colcountx; colc++)
                            {
                                XmlElement xmlelement = appstyleDoc.DocumentElement;
                                XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                                XmlNode stylenode = stylenodelist.Item(0);
                                string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                                string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                                //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                                //int colorBx = Convert.ToInt32(colorB);
                                //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                                string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                                string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                                string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                                string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                                string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                                string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
                                string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                                string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                                string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                                string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                                string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                                string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                                string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                                string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                                string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                                string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                                string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                                string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                                string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                                string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                                int intmergecellcount = Convert.ToInt32(mergecellcount);

                                string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                                string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                                string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                                string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                                string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();

                                /////////////////////////////exception traitement/////////////////////////////////
                                //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                                //xlWorkSheet.Cells.ColumnWidth = 20;
                                //////////////////////////////////////////////////////////////////////////////////

                                Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                                rangeDelx.Font.Name = fontname;
                                rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                                rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
                                //rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                                //rangeDelx.Value2 = fontcolorindex;

                                rangeDelx.Font.Bold = (fontbold == "True");
                                rangeDelx.Font.Italic = (fontitalic == "True");
                                rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                                //rangeDelx.Value2 += "bgcolor" + bgcolorindex;
                                rangeDelx.Interior.Color = Convert.ToInt32(bgcolor);
                                //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                                rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                                rangeDelx.WrapText = (wraptext == "True");
                                rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                                rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);



                                rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                                rangeDelx.RowHeight = Convert.ToDouble(rowheight);



                                ///////////////////////////////////merge process///////////////////////////////////////////
                                if (mergecell == "True")
                                {
                                    if (intmergecellcount > 1)
                                    {
                                        Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc - intmergecellcount + 1], xlWorkSheet.Cells[row, colc]) as Excel.Range;// + intmergecellcount - 1
                                        rangemerge.Merge(false);
                                        //rangemerge.HorizontalAlignment = 1;


                                        //(int countarea = 1; countarea < intmergecellcount; countarea++)
                                        for (int countarea = intmergecellcount - 1; countarea >= 1; countarea--)
                                        {
                                            XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                            int mergecolindex = colc - countarea;//colc + countarea
                                            XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                            XmlNode mergestylenode = mergestylenodelist.Item(0);
                                            mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                            appstyleDoc.Save(stylexml);
                                        }
                                    }
                                }

                                /////////////////////////////////////////////////////////////////////////////////////////
                                mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                                //ne peut pas modifier les cellules fusionner
                                if (mergecellcount == "False")
                                {
                                    rangeDelx.NumberFormat = nomberformat;
                                    try
                                    {
                                        rangeDelx.Locked = (locked == "True");
                                        rangeDelx.Locked = (formulahidden == "True");
                                    }
                                    catch
                                    {
                                    }
                                }
                                ///////////////////////////////////////////////////////////////////////////////////////////

                            }
                        }


                    }
            }
            xlApp.ActiveWindow.DisplayGridlines = false;
            //range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //range.Rows.AutoFit();
            //Excel.Range rangemasquer2 = xlWorkSheet.UsedRange.get_Range("A1", "A14") as Excel.Range;
            //rangemasquer2.EntireRow.Hidden = true;

            //pour consigne de masquage
            //Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            //object[,] values8000 = (object[,])rangeremplace.Value2;
            //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            //{
            //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
            //    if (Regex.Equals(valuedel, "-1"))
            //    {
            //        Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
            //        rangeDely.EntireRow.Hidden = true;
            //    }
            //}

            Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("E1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
            Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
           // Excel.Range rangeDelete3 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count - 1, 1] as Excel.Range;
            //consigne supression
            //rangeDelete.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            //rangeDelete2.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //rangeDelete3.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
            rangeDelete.EntireColumn.Hidden = true;


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            //xlWorkBook.Save();
            if (pathstylerfinal != null) System.IO.Directory.CreateDirectory(pathstylerfinal);
            if (divitylerfinal != null) xlWorkBook.SaveCopyAs(divitylerfinal);
            if (divitylerfinal != null) xlWorkBook.Close(false, misValue, misValue);

            if (divitylerfinal == null) xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

#endregion

        #region methode preface reconstruit le fichier Eval.ptw
        private void NewEval_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook0;//New Eval
            Excel.Workbook xlWorkBook1;// Eval


            Excel.Workbook xlWorkBook2;//Admin.ptw
            Excel.Workbook xlWorkBook3;//Histo.ptw
            Excel.Workbook xlWorkBook4;//Annuel.ptw
            Excel.Workbook xlWorkBook5;//Decis.ptw
            Excel.Workbook xlWorkBook6;//Tres.ptw


            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook1 = xlApp.Workbooks.Open("D:\\ptw\\Eval.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

            xlWorkBook2 = xlApp.Workbooks.Open("D:\\ptw\\Admin.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            xlWorkBook3 = xlApp.Workbooks.Open("D:\\ptw\\Histo.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            xlWorkBook4 = xlApp.Workbooks.Open("D:\\ptw\\Annuel.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            xlWorkBook5 = xlApp.Workbooks.Open("D:\\ptw\\Decis.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            xlWorkBook6 = xlApp.Workbooks.Open("D:\\ptw\\Tres.ptw", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

            xlWorkBook0 = xlApp.Workbooks.Add(misValue);

            xlApp.DisplayAlerts = false;
            xlApp.Calculation = Excel.XlCalculation.xlCalculationManual;


            for (int ns = 1; ns <= xlWorkBook1.Sheets.Count; ns++)
            {
                Excel.Worksheet Elastsheet = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item(xlWorkBook0.Sheets.Count);
                xlWorkBook0.Sheets.Add(misValue, Elastsheet, misValue, misValue);
                Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(ns);
                admin1.Unprotect(misValue);//pour mosaique
                //MessageBox.Show(admin1.Name.ToString());
                Excel.Worksheet admin1X = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item(xlWorkBook0.Sheets.Count);

                admin1X.Name = admin1.Name.ToString();
            }

            Excel.Worksheet sheetn1 = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item("Feuil1");
            Excel.Worksheet sheetn2 = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item("Feuil2");
            Excel.Worksheet sheetn3 = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item("Feuil3");
            sheetn1.Delete();
            sheetn2.Delete();
            sheetn3.Delete();

            for (int i = 3; i <= xlWorkBook1.Names.Count; i++)
            {
                //Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(xlWorkBook2.Sheets.Count);
                //admin1.get_Range("A" + i.ToString(), misValue).Value2 = xlWorkBook2.Names.Item(i, misValue, misValue).Name;
                //admin1.get_Range("B" + i.ToString(), misValue).Value2 = xlWorkBook2.Names.Item(i, misValue, misValue);
                //admin1.get_Range("C" + i.ToString(), misValue).Value2 = xlWorkBook2.Names.Item(i, misValue, misValue).Visible;
                if (xlWorkBook1.Names.Item(i, misValue, misValue).Name.Equals("Zone_d_impression") == false)
                {
                    string champ = xlWorkBook1.Names.Item(i, misValue, misValue).Value.ToString();
                    string namec = xlWorkBook1.Names.Item(i, misValue, misValue).Name.ToString();
                    if (Regex.IsMatch(champ, ":"))
                    {
                        champ.Replace(";", ",");
                    }
                    xlWorkBook0.Names.Add(xlWorkBook1.Names.Item(i, misValue, misValue).Name, champ, true, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                }
                //releaseObject(admin1);
            }
            
            Excel.Worksheet sheetfirst = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item(1);
            sheetfirst.Activate();
            for (int ns = 1; ns <= xlWorkBook1.Sheets.Count; ns++)
            {
                Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(ns);
                Excel.Worksheet admin1X = (Excel.Worksheet)xlWorkBook0.Worksheets.get_Item(ns);
                for (int i = 1; i <= admin1.UsedRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= admin1.UsedRange.Columns.Count; j++)
                    {
                        Excel.Range rangecoller = admin1X.Cells[i, j] as Excel.Range;
                        Excel.Range rangecopy = admin1.Cells[i, j] as Excel.Range;
                        if (rangecopy.Formula != null)
                        {
                            rangecoller.Formula = rangecopy.Formula;
                        }
                    }
                }
            }





            xlWorkBook0.SaveCopyAs("D:\\ptw\\NewEval.xls");
            xlWorkBook1.Close(false, misValue, misValue);
            xlApp.Quit();


            releaseObject(xlWorkBook1);
            releaseObject(xlWorkBook0);
            releaseObject(xlWorkBook2);
            releaseObject(xlWorkBook3);
            releaseObject(xlWorkBook4);
            releaseObject(xlWorkBook5);
            releaseObject(xlWorkBook6);
            releaseObject(xlApp);
        }

        #endregion
        private void newempty(object sender, EventArgs e)
        {

        }
        private void button19_Click(object sender, EventArgs e)
        {//pas3 not worked
            pathstylerfinal = textBox6.Text;
            if (checkBox19.Checked)
            {
                //string[] namestable = { "ACT1", "ACT2", "ACT3", "ACT4", "PAS1", "PAS2", "PAS3", "CR1", "CR2", "CR3", "CR4", "ANN5-1", "ANN5-2", "ANN5-3", "ANN6-1", "ANN6-2", "ANN6-3", "ANN7-1", "ANN7-2", "ANN7-3", "ANN8-1", "ANN8-2", "ANN11-1" };
                string[] namestable = { "ACT1",  "ACT4", "PAS1", "PAS3", "CR1",  "CR3", "ANN5-1", "ANN5-2", "ANN6-1", "ANN6-2", "ANN6-3", "ANN7-1", "ANN7-2", "ANN7-3", "ANN8-1", "ANN8-2", "ANN11-1" };
                int rcont = 1;
                int ccont = 12;

                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";

                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;
                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
                    //model.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Historique");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Historique");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Historique");


                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
                    //set using language to be english
                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E1043");
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B1043");
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range grange = glanguesheet.get_Range("F4", "f1043");
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B1043");
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                   
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //    string strcell = values[rcont, ccont].ToString();

                    int col6000 = 0 ;
                    for (int y = 1; y < userange.Columns.Count; y++)
                    {
                        if (values[userange.Rows.Count, y]!= null)
                        {
                            if (values[userange.Rows.Count, y].ToString() == "6000")
                            {
                                col6000 = y;
                                break;
                            }
                        }
                    }
                    try
                    {
                        
                        for (int x = 1; x < userange.Rows.Count; x++)
                        {
                            if (values[x, col6000] != null)
                            {
                                if (values[x, col6000].ToString() == "9000")
                                {
                                }
                                else
                                { //fr 
                                    copyrange = myworksheet.get_Range(myworksheet.Cells[x, 1], myworksheet.Cells[x, 2]);
                                    copyrange.Copy(misValue);
                                    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                    //en
                                    copyrange = enworksheet.get_Range(enworksheet.Cells[x, 1], enworksheet.Cells[x, 2]);
                                    copyrange.Copy(misValue);
                                    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                    //german
                                    copyrange = gworksheet.get_Range(gworksheet.Cells[x, 1], gworksheet.Cells[x, 2]);
                                    copyrange.Copy(misValue);
                                    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                }
                            }
                            else
                            { //fr 
                                copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[x, 2]);
                                copyrange.Copy(misValue);
                                copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                //en
                                copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[x, 2]);
                                copyrange.Copy(misValue);
                                copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                                //german
                                copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[x, 2]);
                                copyrange.Copy(misValue);
                                copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                            }
                           
                        }


                      

                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}
                   
                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");

                    deletelaguage.Delete();

                    enlanguesheet.Delete();
                    glanguesheet.Delete();


                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();

                    releaseObject(deleteworksheet1);
                    releaseObject(deleteworksheet2);
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);

                    app.Quit();

                }
            }
            if (checkBox22.Checked)
            {
                string[] namestable = { "ANNUEL-CR1", "ANNUEL-CR2", "ANNUEL-CR3", "ANNUEL-BILACT1", "ANNUEL-BILPAS1", "ANNUEL-BILUSACT1", "ANNUEL-BILUSPAS1", "ANNUEL-FLUXFIN1", "ANNUEL-FLUXTRES1", "ANNUEL-RATIOS1", "ANNUEL-RATIOS2", "ANNUEL-SYNTH" };
                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";
                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;

                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Comptes annuels");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Comptes annuels");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Comptes annuels");

                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Admin.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Admin.Langues");
                    //set using language to be english
                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E3511");
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B3511");
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range grange = glanguesheet.get_Range("F4", "F3511");
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B3511");
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    int rcont = 1;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //string strcell = values[rcont, ccont].ToString();

                    try
                    {
                        //fr 
                        copyrange = myworksheet.get_Range(myworksheet.Cells[rcont, 1], myworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //en
                        copyrange = enworksheet.get_Range(enworksheet.Cells[rcont, 1], enworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //german
                        copyrange = gworksheet.get_Range(gworksheet.Cells[rcont, 1], gworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Admin.Langues");
                    deletelaguage.Delete();
                    enlanguesheet.Delete();
                    glanguesheet.Delete();
                    Excel.Worksheet delosheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet2 = (Excel.Worksheet)enworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet3 = (Excel.Worksheet)gworkbook.Worksheets.get_Item("O");
                    delosheet1.Delete();
                    delosheet2.Delete();
                    delosheet3.Delete();
                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    deletesheets(myworkbook);
                    deletesheets(enworkbook);
                    deletesheets(gworkbook);
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);


                    app.Quit();

                }


            }
            else if (checkBox23.Checked)
            {

                string[] namestable = { "EVAL-SYNTHVALO2", "EVAL-SYNTHVALO1", "EVAL-SYNTHMULT1" };
                int rcont = 1;
                int ccont = 12;

                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = "d:\\ptw\\notepme\\" + namestable[i] + "_EN.xlsx";
                    string gpath = "d:\\ptw\\notepme\\" + namestable[i] + "_GER.xlsx";
                    string frpath = "d:\\ptw\\notepme\\" + namestable[i] + "_FR.xlsx";

                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;
                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    // deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    //  deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    // deleteworksheet1.Delete();
                    //deleteworksheet2.Delete();
                    //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
                    //model.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("SynthèseValorisations");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("SynthèseValorisations");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("SynthèseValorisations");


                    //Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
                    //Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
                    ////set using language to be english
                    //Excel.Range enrange = enlanguesheet.get_Range("E4", "E1043");
                    //Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B1043");
                    //enrange.Copy(enpasterange);
                    //releaseObject(enrange);
                    //releaseObject(enpasterange);
                    ////set using language to be german
                    //Excel.Range grange = glanguesheet.get_Range("F4", "f1043");
                    //Excel.Range gpasterange = glanguesheet.get_Range("B4", "B1043");
                    //grange.Copy(gpasterange);
                    //releaseObject(grange);
                    //releaseObject(gpasterange);

                    //Excel.Range userange = myworksheet.UsedRange;
                    //object[,] values = (object[,])userange.Value2;
                    //Excel.Range copyrange;
                    ////for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    ////{
                    ////    string strcell = values[rcont, ccont].ToString();

                    //try
                    //{
                    //    //fr 
                    //    copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[userange.Rows.Count, 2]);
                    //    copyrange.Copy(misValue);
                    //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    //    //en
                    //    copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[userange.Rows.Count, 2]);
                    //    copyrange.Copy(misValue);
                    //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                    //    //german
                    //    copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[userange.Rows.Count, 2]);
                    //    copyrange.Copy(misValue);
                    //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    //}
                    //catch (Exception ex)
                    //{
                    //}
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    //Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");
                    // deletelaguage.Delete();

                    //enlanguesheet.Delete();
                    // glanguesheet.Delete();


                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();

                    //releaseObject(deleteworksheet1);
                    // releaseObject(deleteworksheet2);
                    // releaseObject(deleteensheet);
                    // releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    // releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    //releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);

                    app.Quit();

                }

            }
        }

        private void copyrangeannewlrefer(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
           // prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Range userange = xlWorkSheet.UsedRange;
            object[,] valuex = (object[,])userange.Value2;
            Excel.Range copyrange = xlWorkSheet.UsedRange.get_Range("A35", "A"+userange.Rows.Count) as Excel.Range;
            //copyrange.EntireColumn.Replace("'Comptes annuels'!F", "'Comptes annuels'!D", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);


            for (int col = 4; col < 25; col++)
            {


                Excel.Range psrange = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.UsedRange.Cells[35, col], xlWorkSheet.UsedRange.Cells[100, col]);
                copyrange.Copy(psrange);

            }
            Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Range range = xlWorkSheet1.UsedRange;
            xlWorkSheet1.UsedRange.get_Range("H1", "H848").EntireColumn.Replace("!$N", "!$H", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet1.UsedRange.get_Range("O1", "O848").EntireColumn.Replace("!$N", "!$O", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
            xlWorkSheet1.UsedRange.get_Range("V1", "V848").EntireColumn.Replace("!$N", "!$V", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void deletesheets(Excel.Workbook xlWorkBook)
        {
            Excel.Worksheet Historique = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Worksheet HistPrefac = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet HistCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
            Excel.Worksheet HistLangues = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues");
            Excel.Worksheet HistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-s");
            Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-s");
            Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-s");
            Excel.Worksheet HistRefersheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-s");
            //Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            Excel.Worksheet Paramimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            Excel.Worksheet Paramgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            //Excel.Worksheet AdminLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            Excel.Worksheet AdminServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            Excel.Worksheet Tsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            Excel.Worksheet ParamSavsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            Excel.Worksheet Macrossheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            Excel.Worksheet Vsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            Excel.Worksheet Mosaiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            Excel.Worksheet GraphiquesSRsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            Excel.Worksheet Graphimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            Excel.Worksheet Dontdeletesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            Excel.Worksheet Finsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            Excel.Worksheet ChoixMethodessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            Excel.Worksheet Noterecapitulativesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            Excel.Worksheet SyntheseValorisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Worksheet DefinitionsArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            Excel.Worksheet RappelRetraitementssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            Excel.Worksheet RisqueEntreprisesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            Excel.Worksheet ChoixTauxParamsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            Excel.Worksheet TauxParamArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            Excel.Worksheet CorrectifsSIGBilansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            Excel.Worksheet APNNEsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            Excel.Worksheet FiscaliteDiffereesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            Excel.Worksheet PatrimonialAncAnccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            Excel.Worksheet FondsDeCommercesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            Excel.Worksheet Goodwillsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            Excel.Worksheet AutresCapitalisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            Excel.Worksheet Multiplessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            Excel.Worksheet MethodesMixtessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            Excel.Worksheet TransactionsComparablessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            Excel.Worksheet GordonShapiroBatessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            Excel.Worksheet CalculFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            Excel.Worksheet DiscountedFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            Excel.Worksheet CmpcWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            Excel.Worksheet CmpcWaccArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            Excel.Worksheet ModuleWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            Excel.Worksheet CCEFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            Excel.Worksheet TriRentabiliteProjetsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            Excel.Worksheet TourDeTableSynthesesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            Excel.Worksheet EvalLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            Excel.Worksheet Controlessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            Excel.Worksheet EvalServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            Excel.Worksheet Composantessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            Excel.Worksheet Jsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            Excel.Worksheet Factgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            Excel.Worksheet Lsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            Excel.Worksheet Msheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            Excel.Worksheet Tresoreriesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            Excel.Worksheet ABsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            Excel.Worksheet Paramtresorsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            Excel.Worksheet Saisonnalitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            Excel.Worksheet Zsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
            //coller value de synthese

            Historique.Delete();
            HistPrefac.Delete();
            HistCalculs.Delete();
            HistLangues.Delete();
            HistRefer.Delete();
            Historiquesheet.Delete();
            HistPrefacsheet.Delete();
            HistCalculssheet.Delete();
            HistLanguessheet.Delete();
            HistRefersheet.Delete();
            Identitesheet.Delete();
            Paramimprsheet.Delete();
            Psheet.Delete();
            Paramgenerauxsheet.Delete();
            AdminServicesheet.Delete();
            Tsheet.Delete();
            ParamSavsheet.Delete();
            Macrossheet.Delete();
            Vsheet.Delete();
            Mosaiquesheet.Delete();
            GraphiquesSRsheet.Delete();
            Graphimprsheet.Delete();
            Dontdeletesheet.Delete();
            Finsheet.Delete();
            ChoixMethodessheet.Delete();
            Noterecapitulativesheet.Delete();
            SyntheseValorisationssheet.Delete();
            DefinitionsArrierePlansheet.Delete();
            RappelRetraitementssheet.Delete();
            RisqueEntreprisesheet.Delete();
            ChoixTauxParamsheet.Delete();
            TauxParamArrierePlansheet.Delete();
            CorrectifsSIGBilansheet.Delete();
            APNNEsheet.Delete();
            FiscaliteDiffereesheet.Delete();
            PatrimonialAncAnccsheet.Delete();
            FondsDeCommercesheet.Delete();
            Goodwillsheet.Delete();
            AutresCapitalisationssheet.Delete();
            Multiplessheet.Delete();
            MethodesMixtessheet.Delete();
            TransactionsComparablessheet.Delete();
            GordonShapiroBatessheet.Delete();
            CalculFCFsheet.Delete();
            DiscountedFCFsheet.Delete();
            CmpcWaccsheet.Delete();
            CmpcWaccArrierePlansheet.Delete();
            ModuleWaccsheet.Delete();
            CCEFsheet.Delete();
            TriRentabiliteProjetsheet.Delete();
            TourDeTableSynthesesheet.Delete();
            EvalLanguessheet.Delete();
            Controlessheet.Delete();
            EvalServicesheet.Delete();
            Composantessheet.Delete();
            Jsheet.Delete();
            Factgenerauxsheet.Delete();
            Lsheet.Delete();
            Msheet.Delete();
            Tresoreriesheet.Delete();
            ABsheet.Delete();
            Paramtresorsheet.Delete();
            Saisonnalitesheet.Delete();
            Zsheet.Delete();
            model.Delete();

            xlWorkBook.Save() ;
            //xlApp.Quit();
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                textBox20.AppendText("==> Start Protection des cellules" + System.Environment.NewLine);
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                int time1 = System.Environment.TickCount;
                object misValue = System.Reflection.Missing.Value;
                prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                if (checkBox24.Checked)
                {
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                    LockStateFile(textBox17.Text + "\\lockedStatus.stat", xlWorkSheet);
                }
                if (checkBox25.Checked)
                {
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
                    LockStateFile(textBox17.Text + "\\lockedStatus2.stat", xlWorkSheet);
                }
                int time12 = System.Environment.TickCount;
                int time = ((time12 - time1) / 1000);
                xlWorkBook.Close();
                xlApp.Quit();
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Protection des cellules OK : " + timeto + " s"+System.Environment.NewLine);
                MessageBox.Show("Protection des cellules OK : " + timeto + " s");
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString()+System.Environment.NewLine);
            }
        }

        private string button20tout_Click(object sender, EventArgs e)
        {
             string timeto="";
             try
             {
                 textBox20.AppendText("==> Start Protection des cellules " + System.Environment.NewLine);
                 Excel.Application xlApp;
                 Excel.Workbook xlWorkBook;
                 int time1 = System.Environment.TickCount;
                 object misValue = System.Reflection.Missing.Value;
                 prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
                 xlApp = new Excel.ApplicationClass();
                 xlApp.Visible = true;
                 xlApp.DisplayAlerts = false;
                 xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                 if (checkBox24.Checked)
                 {
                     Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                     LockStateFile(textBox17.Text + "\\lockedStatus.stat", xlWorkSheet);
                 }
                 if (checkBox25.Checked)
                 {
                     Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
                     LockStateFile(textBox17.Text + "\\lockedStatus2.stat", xlWorkSheet);
                 }
                 int time12 = System.Environment.TickCount;
                 int time = ((time12 - time1) / 1000);
                 xlWorkBook.Close();
                 xlApp.Quit();
                 int hours = time / 3600;
                 int minuit = time / 60 - hours * 60;
                 int second = time - minuit * 60 - hours * 3600;
                 timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                 textBox20.AppendText("Protection des cellules OK : " + timeto + " s" + System.Environment.NewLine);
                 // MessageBox.Show("Protection des cellules OK ! : " + timeto + " s");
             }
             catch (Exception ex)
             {
                 textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
             }
            return timeto;
        }
        public void LockStateFile(String filePath, Excel.Worksheet worksheet)
        {
            //alex: get the source file.
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            Excel.Range usedrange = worksheet.UsedRange;
            object[,] values = (object[,])usedrange.Value2;
            int rowsCount = usedrange.Rows.Count;
            int colsCount = usedrange.Columns.Count;

            int i = 1;
            while (i < rowsCount)
            {

                String s = "";
                for (int j = 1; j < colsCount; ++j)
                {
                    if (values[i, j] == null)
                    {

                        s = s + "1";
                    }
                    else
                    {

                        if ((usedrange.Cells[i, j] as Excel.Range).Locked.ToString() == "False")
                        {
                            //alex: if the cell is not locked append text 0 

                            s = s + "0";
                        }
                        else
                        {
                            // alex: else append text 1

                            s = s + "1";
                        }
                    }
                }
                sw.WriteLine(s);
                ++i;
            }
            sw.Close();

        }

        public void LockStateFiles(String filePath, Excel.Worksheet worksheet)
        {
            //alex: get the source file.
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            Excel.Range usedrange = worksheet.UsedRange;
            object[,] values = (object[,])usedrange.Value2;
            int rowsCount = usedrange.Rows.Count;
            int colsCount = usedrange.Columns.Count;

            int i = 1;
            while (i < rowsCount)
            {

                String s = "";
                for (int j = 1; j < colsCount; ++j)
                {
                    if (values[i, j] == null)
                    {

                        s = s + "1";
                    }
                    else
                    {

                        if ((usedrange.Cells[i, j] as Excel.Range).Locked.ToString() == "False")
                        {
                            //alex: if the cell is not locked append text 0 

                            s = s + "0";
                        }
                        else
                        {
                            // alex: else append text 1

                            s = s + "1";
                        }
                    }
                }
                sw.WriteLine(s);
                ++i;
            }
            sw.Close();

        }

        private void button19_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("S'assurer de la création préalable du fichier lockedStatus.stat.", "sure", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    textBox20.AppendText("==> Start Raz PrefaceNP. Le fichier sera sauvé dans D:\\ptw\\notepme !" + System.Environment.NewLine);
                    int timex = System.Environment.TickCount;
                    FileStream file = new FileStream(textBox17.Text + "\\lockedStatus.stat", FileMode.Open, FileAccess.Read);
                    List<String> lockchecklist;
                    lockchecklist = getLockedstatusList(file);
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;

                    object misValue = System.Reflection.Missing.Value;
                    prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
                    xlApp = new Excel.ApplicationClass();
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;
                    xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                    //Excel.Range usedrange = xlWorkSheet.UsedRange;
                    object[,] values = (object[,])xlWorkSheet.UsedRange.Value2;
                    object[,] formulas = (object[,])xlWorkSheet.UsedRange.Formula;
                    int rowcount = xlWorkSheet.UsedRange.Rows.Count - 1;
                    int columncount = xlWorkSheet.UsedRange.Columns.Count;

                    int rowp = 0;
                    int rowp2 = 0;
                    for (int i = 1; i < columncount - 1; i++)
                    {
                        if (values[xlWorkSheet.UsedRange.Rows.Count, i] != null)
                        {
                            if (values[xlWorkSheet.UsedRange.Rows.Count, i].ToString() == "83000")
                            {
                                rowp = i;
                            }
                            if (values[xlWorkSheet.UsedRange.Rows.Count, i].ToString() == "84000")
                            {
                                rowp2 = i;
                                break;
                            }


                        }
                    }
                        for (int i = 1; i < rowcount - 1; i++)
                        {

                            for (int j = 1; j < columncount-2; j++)
                            {
                                if (values[i, j] != null)
                                {
                                    if (j == rowp || j == rowp2)
                                    {
                                        string xss = "ss";
                                    }
                                    if (i == 19 || i == 20 || i == 21)
                                    {
                                    }
                                    else
                                    {
                                        if (getLockedstatus(lockchecklist, i, j) == false)
                                        {

                                            if (values[i, xlWorkSheet.UsedRange.Columns.Count] != null)
                                            {
                                                string v = values[i, xlWorkSheet.UsedRange.Columns.Count].ToString();
                                                if (values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "224000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "242000-12000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "275000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "762000-4000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "763000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "763000-4000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "243000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "243000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "299000-200" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "468000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "471000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "473000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "475000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "478000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "480000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "482000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "485000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "487000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "745000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "745000-3000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "746000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "746000-3000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "768000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "772000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "776000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "780000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-500" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-700" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "814000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "816000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "818000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "308000" || j == rowp || j == rowp2 || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "347000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "353000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "350000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "344000")
                                                {
                                                    values[i, j] = formulas[i, j];
                                                }
                                                else
                                                {
                                                    values[i, j] = 0;
                                                }
                                            }
                                            else
                                            {
                                                //values[i, j] = formulas[i, j];
                                            }
                                        }
                                        else
                                        {
                                            values[i, j] = formulas[i, j];
                                        }
                                    }
                                }
                            }

                        }
                    xlWorkSheet.UsedRange.Value2 = values;

                    xlWorkBook.SaveCopyAs("D:\\ptw\\notepme\\prefaceNP.xlsx");
                    xlWorkBook.Close();
                    xlApp.Quit();
                    int timey = System.Environment.TickCount;
                    int x = (timey - timex) / 1000;
                    int hours = x / 3600;
                    int minuit = x / 60 - hours * 60;
                    int second = x - minuit * 60 - hours * 3600;
                    string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                    textBox20.AppendText("Raz PrefaceNP OK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s" + System.Environment.NewLine);
                    MessageBox.Show("Raz PrefaceNPOK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s");
                }
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString()+System.Environment.NewLine);
            }
        }

        private void button19_Click_1x(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("S'assurer de la création préalable du fichier lockedStatus.stat.", "sure", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    textBox20.AppendText("==> Start Raz PrefaceNP. Le fichier sera sauvé dans D:\\ptw\\notepme !" + System.Environment.NewLine);
                    int timex = System.Environment.TickCount;
                    FileStream file = new FileStream(textBox17.Text + "\\lockedStatus.stat", FileMode.Open, FileAccess.Read);
                    List<String> lockchecklist;
                    lockchecklist = getLockedstatusList(file);
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;

                    object misValue = System.Reflection.Missing.Value;
                    prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
                    xlApp = new Excel.ApplicationClass();
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = false;
                    xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                    //Excel.Range usedrange = xlWorkSheet.UsedRange;
                    object[,] values = (object[,])xlWorkSheet.UsedRange.Value2;
                    object[,] formulas = (object[,])xlWorkSheet.UsedRange.Formula;
                    int rowcount = 633;
                    int columncount = 12;

                    for (int i = 1; i < rowcount - 1; i++)
                    {

                        for (int j = 1; j < columncount - 1; j++)
                        {
                            if (values[i, j] != null)
                            {
                                if (i == 19 || i == 20 || i == 21)
                                {
                                }
                                else
                                {
                                    if (getLockedstatus(lockchecklist, i, j) == false)
                                    {
                                        
                                           
                                                values[i, j] = 0;
                                       
                                    }
                                    else
                                    {
                                        values[i, j] = formulas[i, j];
                                    }
                                }
                            }
                        }

                    }
                    xlWorkSheet.UsedRange.Value2 = values;

                    xlWorkBook.SaveCopyAs("D:\\ptw\\notepme\\prefaceNP.xlsx");
                    xlWorkBook.Close();
                    xlApp.Quit();
                    int timey = System.Environment.TickCount;
                    int x = (timey - timex) / 1000;
                    int hours = x / 3600;
                    int minuit = x / 60 - hours * 60;
                    int second = x - minuit * 60 - hours * 3600;
                    string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                    textBox20.AppendText("Raz PrefaceNP OK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s" + System.Environment.NewLine);
                    MessageBox.Show("Raz PrefaceNPOK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s");
                }
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
        }

        private string button19_Click_tout(object sender, EventArgs e)
        {

            string timeto = "";
            try
            {
                textBox20.AppendText("==> Start Raz PrefaceNP : le fichier sera sauvé dans D:\\ptw\\notepme" + System.Environment.NewLine);
                int timex = System.Environment.TickCount;
                FileStream file = new FileStream(textBox17.Text + "\\lockedStatus.stat", FileMode.Open, FileAccess.Read);
                List<String> lockchecklist;
                lockchecklist = getLockedstatusList(file);
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;

                object misValue = System.Reflection.Missing.Value;
                prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                //Excel.Range usedrange = xlWorkSheet.UsedRange;
                object[,] values = (object[,])xlWorkSheet.UsedRange.Value2;
                object[,] formulas = (object[,])xlWorkSheet.UsedRange.Formula;
                int rowcount = xlWorkSheet.UsedRange.Rows.Count - 1;
                int columncount = xlWorkSheet.UsedRange.Columns.Count;

                for (int i = 1; i < rowcount - 1; i++)
                {

                    for (int j = 1; j < columncount - 1; j++)
                    {
                        if (values[i, j] != null)
                        {
                            if (i == 19 || i == 20 || i == 21)
                            {
                            }
                            else
                            {
                                if (getLockedstatus(lockchecklist, i, j) == false)
                                {
                                    if (values[i, xlWorkSheet.UsedRange.Columns.Count] != null)
                                    {
                                        string v = values[i, xlWorkSheet.UsedRange.Columns.Count].ToString();
                                        if (values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "224000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "242000-12000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "275000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "762000-4000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "763000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "763000-4000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "243000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "243000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "299000-200" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "468000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "471000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "473000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "475000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "478000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "480000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "482000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "485000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "487000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "745000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "745000-3000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "746000-2000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "746000-3000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "768000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "772000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "776000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "780000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-500" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-700" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "791000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "814000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "816000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "818000-1000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "308000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "347000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "353000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "350000" || values[i, xlWorkSheet.UsedRange.Columns.Count].ToString() == "344000")
                                        {
                                            values[i, j] = formulas[i, j];
                                        }
                                        else
                                        {
                                            values[i, j] = 0;
                                        }
                                    }
                                    else
                                    {
                                        //values[i, j] = formulas[i, j];
                                    }
                                }
                                else
                                {
                                    values[i, j] = formulas[i, j];
                                }
                            }
                        }
                    }

                }
                xlWorkSheet.UsedRange.Value2 = values;

                xlWorkBook.SaveCopyAs("D:\\ptw\\notepme\\prefaceNP.xlsx");
                xlWorkBook.Close();
                xlApp.Quit();
                int timey = System.Environment.TickCount;
                int x = (timey - timex) / 1000;
                int hours = x / 3600;
                int minuit = x / 60 - hours * 60;
                int second = x - minuit * 60 - hours * 3600;
                timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Raz PrefaceNP OK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s" + System.Environment.NewLine);
                //MessageBox.Show("Raz PrefaceNP\nOK ! Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s");

            }
            catch(Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
                return timeto;
        }

        public List<String> getLockedstatusList(FileStream f)
        {
            // Création d'un objet  "sr" de type "StreamReader".
            StreamReader sr = new StreamReader(f);
            // Création de la liste de chaine de caractère initiale.
            List<String> lines = new List<String>();
            String line;

            while ((line = sr.ReadLine()) != null)
            //Tant qu'il existe des ligne dans le fichier.
            {
                // Ajout de chaque ligne à la liste.
                lines.Add(line);

            }
            // fermeture de l'objet "sr".
            sr.Close();
            // fermeture du fichier "lockedStatus.stat" pour libérer sa place mémoire. 
            f.Close();
            return lines;
        }

        public bool getLockedstatus(List<String> listLocked, int row, int column)
        {
            // Définition de la chaine de caractères correspondant à la ligne demandée (row - 1 -->  la liste "listlocked" commence par l'indexe 0
            // alors que les indexes d'une feuille de calcule Excel commence par 1. Même remarque pour les colonnes column - 1)
            String line = listLocked[row - 1];

            if (line[column - 1] == '0')
            {
                // Si le caractère défini par le numéro "column - 1" est égale à 0 alors,
                // la cellule n'est pas protégée.
                return false;
            }
            else
            {
                // sinon elle n'est pas protégée.
                return true;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialogx = new FolderBrowserDialog();
            FolderBrowserDialogx.RootFolder = Environment.SpecialFolder.MyComputer;
            FolderBrowserDialogx.ShowDialog();
            textBox17.Text = FolderBrowserDialogx.SelectedPath.ToString();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialogx = new OpenFileDialog();
            OpenFileDialogx.InitialDirectory = "D:\\ptw\\";
            OpenFileDialogx.Filter = "Excel Files .xlsx|*.xlsx|ptw files .ptw|*.ptw|All files (*.*)|*.*";
            OpenFileDialogx.RestoreDirectory = true;
            OpenFileDialogx.ShowDialog();
            textBox18.Text = OpenFileDialogx.FileName;
        }

        //private void button23_Click(object sender, EventArgs e)
        //{
        //    Excel.Application xlApp=new Excel.ApplicationClass();
        //    Excel.Workbook xlWorkBook;
        //    xlApp.DisplayAlerts = false;
        //    xlApp.Visible = true;
        //    string path = "D:\\ptw\\temp\\";

        //   string[] strFiles = Directory.GetFiles(path);
        //   foreach(string strfile in strFiles)
        //   {
        //       path += strfile;
        //       xlWorkBook = xlApp.Workbooks.Open(strfile,0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

        //       Excel.Worksheet xlworksheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
        //       Excel.Range hide1 = xlworksheet.get_Range("Y1",xlworksheet.Cells[1,xlworksheet.UsedRange.Columns.Count]);
        //       hide1.Columns.Hidden = true;
        //       Excel.Range hide2 = xlworksheet.get_Range(xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1], xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1]);
        //       hide2.EntireRow.Hidden = true;
        //       xlWorkBook.Save();
        //       xlWorkBook.Close();
   
        //   }
        //   xlApp.Quit();
          

        //}

        private void formarstyle(string path)
        {
            ////////////////open excel////////////////////////
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp2 = new Excel.ApplicationClass();
            xlApp2.Visible = true;
            xlApp2.DisplayAlerts = false;
            string openfilex = path;
            xlWorkBook = xlApp2.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[1];
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////

            if (textBox10.Text != null)
                stylexml = textBox10.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");

            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);

            /////////////////////////////////////set palette couleur///////////////////////////
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);

            range.EntireRow.Font.Size = 8;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;
            int col15000 = 0;
            int colannuel9000 = 0;
            rCnt = range.Rows.Count;

            int col83000 = 0;
            int col8000 = 0;

            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
            col15000 = cf.FindCodedColumn("6000", range);
            //colannuel9000 = cf.FindCodedColumn("9000-1000", range);
            col8000 = cf.FindCodedColumn("8000", range);
            col83000 = cf.FindCodedColumn("83000", range);

            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint + 1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////


            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;
            int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
            for (row = 1; row <= rowCountx - 1; row++)
            {
                string value = Convert.ToString(values[row, col15000]);
                for (int nbs = 1; nbs <= nbtotalint; nbs++)
                {
                    if (Regex.Equals(value, tablestyle[nbs]))
                    {
                        XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                        if (xstyle != null)
                        {
                            colcount = (xstyle.SelectSingleNode("col")).InnerText;
                        }
                        int colcountx = Convert.ToInt32(colcount);
                        for (int colc = 1; colc <= colcountx; colc++)
                        {
                            XmlElement xmlelement = appstyleDoc.DocumentElement;
                            XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                            XmlNode stylenode = stylenodelist.Item(0);
                            string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                            string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                            //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                            //int colorBx = Convert.ToInt32(colorB);
                            //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                            string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                            int fcolor = Convert.ToInt32(fontcolor);
                            //string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                            string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                            string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                            string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                            string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
                            //int bcolor = Convert.ToInt32(bgcolor);
                            string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                            string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                            string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                            string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                            string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                            string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                            string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                            string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                            string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                            string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                            string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                            string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                            string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                            string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            int intmergecellcount = Convert.ToInt32(mergecellcount);

                            string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                            string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                            string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                            string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                            string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
                            ///////////////////////////////////merge process///////////////////////////////////////////
                            if (mergecell == "True")
                            {
                                if (intmergecellcount > 1)
                                {
                                    Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
                                    rangemerge.Merge(false);
                                    //rangemerge.HorizontalAlignment = 1;

                                    for (int countarea = 1; countarea < intmergecellcount; countarea++)
                                    {
                                        XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                        int mergecolindex = colc + countarea;
                                        XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                        XmlNode mergestylenode = mergestylenodelist.Item(0);
                                        mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                        appstyleDoc.Save(stylexml);
                                    }
                                }
                            }
                            /////////////////////////////exception traitement/////////////////////////////////
                            //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                            //xlWorkSheet.Cells.ColumnWidth = 20;
                            //////////////////////////////////////////////////////////////////////////////////

                            /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
                            Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                            rangeDelx.Font.Name = fontname;
                            rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                            // rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                            rangeDelx.Font.Color = fcolor;


                            rangeDelx.Font.Bold = (fontbold == "True");
                            rangeDelx.Font.Italic = (fontitalic == "True");
                            rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                            //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);
                            rangeDelx.Interior.Color = bgcolor;
                            // rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                            rangeDelx.WrapText = (wraptext == "True");
                            rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                            rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

                            /////////////////////////////////////////////////////////////////////////////////////////
                            mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            //ne peut pas modifier les cellules fusionner
                            if (mergecellcount == "1")
                            {
                                rangeDelx.NumberFormat = nomberformat;
                                try
                                {
                                    rangeDelx.Locked = (locked == "True");
                                    rangeDelx.Locked = (formulahidden == "True");
                                }
                                catch
                                {
                                }
                            }
                            ///////////////////////////////////////////////////////////////////////////////////////////
                            rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                            rangeDelx.RowHeight = Convert.ToDouble(rowheight);
                        }
                    }
                }
            }
            xlApp2.ActiveWindow.DisplayGridlines = false;
            //pour consigne de masquage
            Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            object[,] values8000 = (object[,])rangeremplace.Value2;
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                if (Regex.Equals(valuedel, "-1"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col83000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                if (Regex.Equals(valuedel, "-5"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col8000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");
            //xlWorkBook.Save();
            xlWorkBook.Save();
           
            xlApp2.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp2);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try
            {

                textBox20.AppendText("==> Start Formatage des styles de HISTORIQUE dans PrefaceNP " + System.Environment.NewLine);

                int time1 = System.Environment.TickCount;
                if (File.Exists("D:\\ptw\\notepme\\prefaceNP.xlsx"))
                {
                    formarstyle("D:\\ptw\\notepme\\prefaceNP.xlsx");
                }
                else if (File.Exists("D:\\ptw\\notepme\\prefaceNPs.xlsx"))
                {
                    formarstyle("D:\\ptw\\notepme\\prefaceNPs.xlsx");
                }
                
                //object misValue = System.Reflection.Missing.Value;
                //Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
                //xlapp.DisplayAlerts = false;
                //xlapp.Visible = true;

                //Excel.Workbook xlworkbookStyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");
                //Excel.Workbook xlworkbookNP = xlapp.Workbooks.Open("D:\\ptw\\notepme\\prefaceNP.xlsx");

                //Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookStyle.Worksheets.get_Item("Histo et Histo-s");
                //Excel._Worksheet xlworksheet = (Excel.Worksheet)xlworkbookNP.Worksheets.get_Item("Historique");

                ////Excel.Range range= xlworksheet.get_Range("A15","M1856");
                //Excel.Range rangeAll = xlworksheet.UsedRange;
                //rangeAll.ClearFormats();
                //rangeAll.Interior.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 0, 0));

                ////xlworkbookNP.SaveAs("D:\\ptw\\changeStyle\\123.xlsx");
                //object[,] values1 = (object[,])rangeAll.Value2;
                ////int col6000;
                ////for (int i = 1; i < rangeAll.Columns.Count;i++)
                ////{
                ////    if (values1[rangeAll.Rows.Count, i].ToString() != null)
                ////    {
                ////        if (values1[rangeAll.Rows.Count, i].ToString() == "6000")
                ////        {
                ////            col6000 = i;
                ////            break;
                ////        }
                ////    }
                ////}
                //Excel.Range rangeUsedStyle = xlworksheetStyle.UsedRange;
                //int col17000style = 0;
                //object[,] valueStyle = (object[,])rangeUsedStyle.Value2;
                //for (int i = 1; i <= rangeUsedStyle.Columns.Count; i++)
                //{
                //    if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() != null)
                //    {
                //        if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() == "17000")
                //        {
                //            col17000style = i;
                //            break;
                //        }
                //    }
                //}


                //for (int i = 14; i <= 79; i++)
                //{
                //    Excel.Range chercheStyle = xlworksheetStyle.get_Range("A" + i, "A" + i);
                //    Excel.Range changeFontRange = xlworksheetStyle.Cells[i, col17000style] as Excel.Range;
                //    if (changeFontRange.Value2 != null)
                //    {
                //        int fontToChange = int.Parse(changeFontRange.Value2.ToString());

                //        Excel.Range rangeToChangeFont = xlworksheetStyle.get_Range("C" + i, "O" + i);
                //        rangeToChangeFont.EntireRow.Font.Size = fontToChange;




                //    }
                //    string cherche = "";
                //    if (chercheStyle.Value2 != null)
                //    {
                //        cherche = chercheStyle.Value2.ToString();
                //    }
                //    else
                //    {
                //        continue;
                //    }
                //    Excel.Range rangeStyle = xlworksheetStyle.get_Range("C" + i, "O" + i);
                //    rangeStyle.Copy();

                //    int row8180002000 = 0;
                //    for (int m = 1; m <= rangeAll.Rows.Count; m++)
                //    {
                //        if (values1[m, rangeAll.Columns.Count] != null)
                //        {
                //            if (values1[m, rangeAll.Columns.Count].ToString() == "818000-2000")
                //            {
                //                row8180002000 = m;
                //                break;

                //            }
                //        }
                //    }
                //    for (int t = 13; t <= row8180002000; t++)
                //    {

                //        Excel.Range rangeColN = xlworksheet.get_Range("N" + t, "N" + t);
                //        if (rangeColN.Value2 != null)
                //        {
                //            string x = rangeColN.Value2.ToString();
                //            if (cherche == "12000")
                //            {
                //                int sd = 0;
                //                string xss = rangeColN.Value2.ToString();
                //            }
                //            if (rangeColN.Value2.ToString() == cherche)
                //            {
                               
                //                Excel.Range rangePasteStyle = xlworksheet.get_Range("A" + t, "M" + t);
                //                rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                //                rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                            
                //                //if (changeFontRange.Value2 != null)
                //                //{
                //                //    int fontToChange = int.Parse(changeFontRange.Value2.ToString());
                //                //    rangePasteStyle.EntireRow.Font.Size = fontToChange;
                //                //}
                //                if (rangeColN.Value2.ToString() == "4000-100" || rangeColN.Value2.ToString() == "12000-1000" || rangeColN.Value2.ToString() == "10000-100" || rangeColN.Value2.ToString() == "10000-200" || rangeColN.Value2.ToString() == "10000")
                //                {
                //                    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                //                    rangeAutoFit.EntireRow.RowHeight = 7; ;
                //                }//------------------------hide some lines--------------------------------------------
                //                else if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
                //                {
                //                    Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                //                    rangeToHide.Hidden = true;
                //                }
                //                //------------------------------------------------------------------------------------


                //                    //------------------------auto fit the height of the big title rows------------------------
                //                else if (rangeColN.Value2.ToString() == "6000" || rangeColN.Value2.ToString() == "5000")
                //                {
                //                    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                //                    rangeAutoFit.EntireRow.AutoFit();
                //                }
                //            }
                //            if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000" )
                //            {
                //                Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                //                rangeToHide.Hidden = true;
                //            }
                //            //else if (rangeColN.Value2.ToString() == "4000-100" || rangeColN.Value2.ToString() == "12000-1000" || rangeColN.Value2.ToString() == "10000-100" || rangeColN.Value2.ToString() == "10000-200" || rangeColN.Value2.ToString() == "10000")
                //            //{
                //            //    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                //            //    rangeAutoFit.EntireRow.RowHeight = 7; ;
                //            //}
                //        }
                //    }
                //}


                //xlworkbookNP.Save();
                //xlapp.Quit();
                int time = System.Environment.TickCount;
                time = time - time1;
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s" + System.Environment.NewLine);
                MessageBox.Show("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s");

            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString());
            }
        }
        private void button23_Clicks(object sender, EventArgs e)
        {
            try{
                textBox20.AppendText("==> Start Formatage des styles de HISTORIQUE dans PrefaceNP " + System.Environment.NewLine);


                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
                xlapp.DisplayAlerts = false;
                xlapp.Visible = true;

                int time1 = System.Environment.TickCount;
                Excel.Workbook xlworkbookStyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");
                Excel.Workbook xlworkbookNP = xlapp.Workbooks.Open("D:\\ptw\\notepme\\prefaceNPS.xlsx");

                Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookStyle.Worksheets.get_Item("Histo et Histo-s");
                Excel._Worksheet xlworksheet = (Excel.Worksheet)xlworkbookNP.Worksheets.get_Item("Historique");

                //Excel.Range range= xlworksheet.get_Range("A15","M1856");
                Excel.Range rangeAll = xlworksheet.get_Range("A1","BL661");;
                rangeAll.ClearFormats();
                rangeAll.Interior.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 0, 0));

                //xlworkbookNP.SaveAs("D:\\ptw\\changeStyle\\123.xlsx");
                object[,] values1 = (object[,])rangeAll.Value2;
                //int col6000;
                //for (int i = 1; i < rangeAll.Columns.Count;i++)
                //{
                //    if (values1[rangeAll.Rows.Count, i].ToString() != null)
                //    {
                //        if (values1[rangeAll.Rows.Count, i].ToString() == "6000")
                //        {
                //            col6000 = i;
                //            break;
                //        }
                //    }
                //}
                Excel.Range rangeUsedStyle = xlworksheetStyle.UsedRange;
                int col17000style = 0;
                object[,] valueStyle = (object[,])rangeUsedStyle.Value2;
                for (int i = 1; i <= rangeUsedStyle.Columns.Count; i++)
                {
                    if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() != null)
                    {
                        if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() == "17000")
                        {
                            col17000style = i;
                            break;
                        }
                    }
                }


                for (int i = 14; i <= 79; i++)
                {
                    Excel.Range chercheStyle = xlworksheetStyle.get_Range("A" + i, "A" + i);
                    Excel.Range changeFontRange = xlworksheetStyle.Cells[i, col17000style] as Excel.Range;
                    if (changeFontRange.Value2 != null)
                    {
                        int fontToChange = int.Parse(changeFontRange.Value2.ToString());

                        Excel.Range rangeToChangeFont = xlworksheetStyle.get_Range("C" + i, "O" + i);
                        rangeToChangeFont.Font.Size = fontToChange;




                    }
                    string cherche = "";
                    if (chercheStyle.Value2 != null)
                    {
                        cherche = chercheStyle.Value2.ToString();
                    }
                    else
                    {
                        continue;
                    }
                    Excel.Range rangeStyle = xlworksheetStyle.get_Range("C" + i, "O" + i);
                    rangeStyle.Copy();

                    int row8180002000 = 0;
                    for (int m = 1; m <= rangeAll.Rows.Count; m++)
                    {
                        if (values1[m, rangeAll.Columns.Count] != null)
                        {
                            if (values1[m, rangeAll.Columns.Count].ToString() == "380000-91000")
                            {
                                row8180002000 = m;
                                break;

                            }
                        }
                    }
                    for (int t = 13; t <= row8180002000; t++)
                    {

                        Excel.Range rangeColN = xlworksheet.get_Range("N" + t, "N" + t);
                        if (rangeColN.Value2 != null)
                        {
                            if (rangeColN.Value2.ToString() == cherche)
                            {

                                Excel.Range rangePasteStyle = xlworksheet.get_Range("A" + t, "M" + t);
                                rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                //------------------------hide some lines--------------------------------------------
                                if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
                                {
                                    Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                    rangeToHide.Hidden = true;
                                }
                                //------------------------------------------------------------------------------------


                                   //------------------------auto fit the height of the big title rows------------------------
                                else if (rangeColN.Value2.ToString() == "6000" || rangeColN.Value2.ToString() == "5000")
                                {
                                    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                    rangeAutoFit.EntireRow.AutoFit();
                                }
                            }
                            else if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
                            {
                                Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                rangeToHide.Hidden = true;
                            }

                        }
                    }
                }


                xlworkbookNP.Save();
                xlapp.Quit();
                int time = System.Environment.TickCount;
                time = time - time1;
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s" + System.Environment.NewLine);
                MessageBox.Show("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s");
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString());
            }

        }
        // alex the same previews one
        private string button23_Clicktout(object sender, EventArgs e)
        {

            string timeto = "";
            int timex = System.Environment.TickCount;
            try
            {

                textBox20.AppendText("==> Start Formatage des styles de HISTORIQUE dans PrefaceNP " + System.Environment.NewLine);


                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
                xlapp.DisplayAlerts = false;
                xlapp.Visible = true;

                int time1 = System.Environment.TickCount;
                Excel.Workbook xlworkbookStyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");
                Excel.Workbook xlworkbookNP = xlapp.Workbooks.Open("D:\\ptw\\notepme\\prefaceNP.xlsx");

                Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookStyle.Worksheets.get_Item("Histo et Histo-s");
                Excel._Worksheet xlworksheet = (Excel.Worksheet)xlworkbookNP.Worksheets.get_Item("Historique");

                //Excel.Range range= xlworksheet.get_Range("A15","M1856");
                Excel.Range rangeAll = xlworksheet.UsedRange;
                rangeAll.ClearFormats();
                rangeAll.Interior.ColorIndex = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 0, 0));

                //xlworkbookNP.SaveAs("D:\\ptw\\changeStyle\\123.xlsx");
                object[,] values1 = (object[,])rangeAll.Value2;
                //int col6000;
                //for (int i = 1; i < rangeAll.Columns.Count;i++)
                //{
                //    if (values1[rangeAll.Rows.Count, i].ToString() != null)
                //    {
                //        if (values1[rangeAll.Rows.Count, i].ToString() == "6000")
                //        {
                //            col6000 = i;
                //            break;
                //        }
                //    }
                //}
                Excel.Range rangeUsedStyle = xlworksheetStyle.UsedRange;
                int col17000style = 0;
                object[,] valueStyle = (object[,])rangeUsedStyle.Value2;
                for (int i = 1; i <= rangeUsedStyle.Columns.Count; i++)
                {
                    if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() != null)
                    {
                        if (valueStyle[rangeUsedStyle.Rows.Count, i].ToString() == "17000")
                        {
                            col17000style = i;
                            break;
                        }
                    }
                }


                for (int i = 14; i <= 79; i++)
                {
                    Excel.Range chercheStyle = xlworksheetStyle.get_Range("A" + i, "A" + i);
                    Excel.Range changeFontRange = xlworksheetStyle.Cells[i, col17000style] as Excel.Range;
                    if (changeFontRange.Value2 != null)
                    {
                        int fontToChange = int.Parse(changeFontRange.Value2.ToString());

                        Excel.Range rangeToChangeFont = xlworksheetStyle.get_Range("C" + i, "O" + i);
                        rangeToChangeFont.EntireRow.Font.Size = fontToChange;




                    }
                    string cherche = "";
                    if (chercheStyle.Value2 != null)
                    {
                        cherche = chercheStyle.Value2.ToString();
                    }
                    else
                    {
                        continue;
                    }
                    Excel.Range rangeStyle = xlworksheetStyle.get_Range("C" + i, "O" + i);
                    rangeStyle.Copy();

                    int row8180002000 = 0;
                    for (int m = 1; m <= rangeAll.Rows.Count; m++)
                    {
                        if (values1[m, rangeAll.Columns.Count] != null)
                        {
                            if (values1[m, rangeAll.Columns.Count].ToString() == "818000-2000")
                            {
                                row8180002000 = m;
                                break;

                            }
                        }
                    }
                    for (int t = 13; t <= row8180002000; t++)
                    {

                        Excel.Range rangeColN = xlworksheet.get_Range("N" + t, "N" + t);
                        if (rangeColN.Value2 != null)
                        {
                            string x = rangeColN.Value2.ToString();
                            if (cherche == "12000")
                            {
                                int sd = 0;
                                string xss = rangeColN.Value2.ToString();
                            }
                            if (rangeColN.Value2.ToString() == cherche)
                            {
                               
                                Excel.Range rangePasteStyle = xlworksheet.get_Range("A" + t, "M" + t);
                                rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                rangePasteStyle.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                            
                                //if (changeFontRange.Value2 != null)
                                //{
                                //    int fontToChange = int.Parse(changeFontRange.Value2.ToString());
                                //    rangePasteStyle.EntireRow.Font.Size = fontToChange;
                                //}
                                if (rangeColN.Value2.ToString() == "4000-100" || rangeColN.Value2.ToString() == "12000-1000" || rangeColN.Value2.ToString() == "10000-100" || rangeColN.Value2.ToString() == "10000-200" || rangeColN.Value2.ToString() == "10000")
                                {
                                    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                    rangeAutoFit.EntireRow.RowHeight = 7; ;
                                }//------------------------hide some lines--------------------------------------------
                                else if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000")
                                {
                                    Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                    rangeToHide.Hidden = true;
                                }
                                //------------------------------------------------------------------------------------


                                    //------------------------auto fit the height of the big title rows------------------------
                                else if (rangeColN.Value2.ToString() == "6000" || rangeColN.Value2.ToString() == "5000")
                                {
                                    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                    rangeAutoFit.EntireRow.AutoFit();
                                }
                            }
                            if (rangeColN.Value2.ToString() == "0" || rangeColN.Value2.ToString() == "9000" )
                            {
                                Excel.Range rangeToHide = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                                rangeToHide.Hidden = true;
                            }
                            //else if (rangeColN.Value2.ToString() == "4000-100" || rangeColN.Value2.ToString() == "12000-1000" || rangeColN.Value2.ToString() == "10000-100" || rangeColN.Value2.ToString() == "10000-200" || rangeColN.Value2.ToString() == "10000")
                            //{
                            //    Excel.Range rangeAutoFit = xlworksheet.get_Range("A" + t, "A" + t).EntireRow;
                            //    rangeAutoFit.EntireRow.RowHeight = 7; ;
                            //}
                        }
                    }
                }


                xlworkbookNP.Save();
                xlapp.Quit();
                int time = System.Environment.TickCount;
                time = time - time1;
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                 timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s" + System.Environment.NewLine);
                //MessageBox.Show("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s");

            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString());
            }
            return timeto;
        }
        private void formartAnuel(string path)
        {
            ////////////////open excel////////////////////////
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;


            string openfilex = path;
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Feuil1");
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Comptes annuels"];
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            //////////////////////////////open le fichier style XML//////////////////////

            if (textBox10.Text != null)
                stylexml = textBox10.Text;
            else
                MessageBox.Show("Veuillez choiser le fichier style en format XML");


            XmlDocument appstyleDoc = new XmlDocument();
            appstyleDoc.Load(stylexml);

            /////////////////////////////////////set palette couleur///////////////////////////
            xlWorkBook.ResetColors();
            XmlElement indexxmlelement = appstyleDoc.DocumentElement;
            XmlNodeList indexstylenodelist = indexxmlelement.SelectNodes("//palette");
            XmlNode indexstylenode = indexstylenodelist.Item(0);
            range.EntireRow.Font.Size = 8;
            ////////////////////////////////////////////////////////////////////////////////////

            int rCnt = 0;
            int cCnt = 0;
            //int col = 0;
            int col15000 = 0;
            int col11000 = 0;
            rCnt = range.Rows.Count;


            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "9000-1000"))
                {
                    col15000 = cCnt;

                }
                if (Regex.Equals(valuecellabs, "11000-1000"))
                {
                    col11000 = cCnt;
                    break;
                }
            }

            ///////////////////////////////////construit tableaux style///////////MAX 100///////////////
            XmlElement nbstyle = appstyleDoc.DocumentElement;
            XmlNodeList nbstylelist = indexxmlelement.SelectNodes("//nbstyle");
            XmlNode nbstylenode = nbstylelist.Item(0);

            string nbtotal = nbstylenode.Attributes["NB"].InnerText.ToString();
            int nbtotalint = 120;
            nbtotalint = Convert.ToInt32(nbtotal);
            string[] tablestyle = new string[nbtotalint + 1];
            for (int nbs = 1; nbs <= nbtotalint; nbs++)
            {
                tablestyle[nbs] = nbstylenode.SelectNodes("nbstyle" + nbs).Item(0).InnerText.ToString();
            }
            /////////////////////////////////////////////////////////////////////////////////////////////


            int row = 1;
            string colcount = "";
            int time1 = System.Environment.TickCount;
            int rowCountx = xlWorkSheet.UsedRange.Rows.Count;
            for (row = 1; row <= rowCountx - 1; row++)
            {
                string value = Convert.ToString(values[row, col15000]);
                for (int nbs = 1; nbs <= nbtotalint; nbs++)
                {
                    if (Regex.Equals(value, tablestyle[nbs]))
                    {
                        XmlNode xstyle = appstyleDoc.SelectSingleNode("//style" + tablestyle[nbs]);
                        if (xstyle != null)
                        {
                            colcount = (xstyle.SelectSingleNode("col")).InnerText;
                        }
                        int colcountx = Convert.ToInt32(colcount);
                        for (int colc = 1; colc <= colcountx; colc++)
                        {
                            XmlElement xmlelement = appstyleDoc.DocumentElement;
                            XmlNodeList stylenodelist = xmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + colc);
                            XmlNode stylenode = stylenodelist.Item(0);
                            string fontname = stylenode.SelectNodes("font").Item(0).InnerText.ToString();
                            string fontsize = stylenode.SelectNodes("fontsize").Item(0).InnerText.ToString();
                            //string colorR = stylenode.SelectNodes("fontcolor").Item(0).Attributes["R"].InnerText.ToString();
                            //int colorBx = Convert.ToInt32(colorB);
                            //int fontcolor = (colorBx * 65536) + (colorGx * 256) + colorRx;
                            string fontcolor = stylenode.SelectNodes("fontcolor").Item(0).InnerText.ToString();
                            // string fontcolorindex = stylenode.SelectNodes("fontcolorindex").Item(0).InnerText.ToString();

                            string fontbold = stylenode.SelectNodes("fontbold").Item(0).InnerText.ToString();
                            string fontitalic = stylenode.SelectNodes("fontitalic").Item(0).InnerText.ToString();
                            string fontunderline = stylenode.SelectNodes("fontunderline").Item(0).InnerText.ToString();

                            string bgcolor = stylenode.SelectNodes("bgcolor").Item(0).InnerText.ToString();
                            string bgcolorindex = stylenode.SelectNodes("bgcolorindex").Item(0).InnerText.ToString();
                            string bordertop = stylenode.SelectNodes("bordertop").Item(0).InnerText.ToString();
                            string borderbot = stylenode.SelectNodes("borderbot").Item(0).InnerText.ToString();
                            string borderleft = stylenode.SelectNodes("borderleft").Item(0).InnerText.ToString();
                            string borderright = stylenode.SelectNodes("borderright").Item(0).InnerText.ToString();
                            string borderweighttop = stylenode.SelectNodes("borderweighttop").Item(0).InnerText.ToString();
                            string borderweightbot = stylenode.SelectNodes("borderweightbot").Item(0).InnerText.ToString();
                            string borderweightleft = stylenode.SelectNodes("borderweightleft").Item(0).InnerText.ToString();
                            string borderweightright = stylenode.SelectNodes("borderweightright").Item(0).InnerText.ToString();

                            string wraptext = stylenode.SelectNodes("wraptext").Item(0).InnerText.ToString();
                            string Halignment = stylenode.SelectNodes("Halignment").Item(0).InnerText.ToString();
                            string Valignment = stylenode.SelectNodes("Valignment").Item(0).InnerText.ToString();
                            string mergecell = stylenode.SelectNodes("mergecell").Item(0).InnerText.ToString();
                            string mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            int intmergecellcount = Convert.ToInt32(mergecellcount);

                            string nomberformat = stylenode.SelectNodes("nomberformat").Item(0).InnerText.ToString();
                            string locked = stylenode.SelectNodes("locked").Item(0).InnerText.ToString();
                            string formulahidden = stylenode.SelectNodes("formulahidden").Item(0).InnerText.ToString();
                            string colwidth = stylenode.SelectNodes("colwidth").Item(0).InnerText.ToString();
                            string rowheight = stylenode.SelectNodes("rowheight").Item(0).InnerText.ToString();
                            ///////////////////////////////////merge process///////////////////////////////////////////
                            if (mergecell == "True")
                            {
                                if (intmergecellcount > 1)
                                {
                                    Excel.Range rangemerge = xlWorkSheet.UsedRange.get_Range(xlWorkSheet.Cells[row, colc], xlWorkSheet.Cells[row, colc + intmergecellcount - 1]) as Excel.Range;
                                    rangemerge.Merge(false);
                                    //rangemerge.HorizontalAlignment = 1;

                                    for (int countarea = 1; countarea < intmergecellcount; countarea++)
                                    {
                                        XmlElement mergexmlelement = appstyleDoc.DocumentElement;
                                        int mergecolindex = colc + countarea;
                                        XmlNodeList mergestylenodelist = mergexmlelement.SelectNodes("//style" + tablestyle[nbs] + "." + mergecolindex);
                                        XmlNode mergestylenode = mergestylenodelist.Item(0);
                                        mergestylenode.SelectNodes("mergecell").Item(0).InnerText = "False";
                                        appstyleDoc.Save(stylexml);
                                    }
                                }
                            }
                            /////////////////////////////exception traitement/////////////////////////////////
                            //Excel.Range rangeLarge = xlWorkSheet.UsedRange as Excel.Range;
                            //xlWorkSheet.Cells.ColumnWidth = 20;
                            //////////////////////////////////////////////////////////////////////////////////

                            /////////////////////////////////////appliquer sur fichier EXCEL//////////////////////////////
                            Excel.Range rangeDelx = xlWorkSheet.Cells[row, colc] as Excel.Range;
                            rangeDelx.Font.Name = fontname;
                            rangeDelx.Font.Size = Convert.ToInt32(fontsize);
                            //2003-2010
                            rangeDelx.Font.Color = Convert.ToInt32(fontcolor);
                            //rangeDelx.Font.ColorIndex = Convert.ToInt32(fontcolorindex);
                            //rangeDelx.Value2 = fontcolorindex;

                            rangeDelx.Font.Bold = (fontbold == "True");
                            rangeDelx.Font.Italic = (fontitalic == "True");
                            rangeDelx.Font.Underline = Convert.ToInt32(fontunderline);
                            //rangeDelx.Value2 += "bgcolor" + bgcolorindex;
                            rangeDelx.Interior.Color = Convert.ToInt32(bgcolor);
                            //rangeDelx.Interior.ColorIndex = Convert.ToInt32(bgcolorindex);

                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Convert.ToInt32(borderweighttop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Convert.ToInt32(bordertop);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Convert.ToInt32(borderweightbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Convert.ToInt32(borderbot);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Convert.ToInt32(borderweightleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Convert.ToInt32(borderleft);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Convert.ToInt32(borderweightright);
                            rangeDelx.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Convert.ToInt32(borderright);

                            rangeDelx.WrapText = (wraptext == "True");
                            rangeDelx.HorizontalAlignment = Convert.ToInt32(Halignment);
                            rangeDelx.VerticalAlignment = Convert.ToInt32(Valignment);

                            /////////////////////////////////////////////////////////////////////////////////////////
                            mergecellcount = stylenode.SelectNodes("mergecellcount").Item(0).InnerText.ToString();
                            //ne peut pas modifier les cellules fusionner
                            if (mergecellcount == "1")
                            {
                                try
                                {
                                    rangeDelx.NumberFormat = nomberformat;
                                    rangeDelx.Locked = (locked == "True");
                                    rangeDelx.Locked = (formulahidden == "True");
                                }
                                catch
                                {
                                }
                            }
                            ///////////////////////////////////////////////////////////////////////////////////////////
                            rangeDelx.ColumnWidth = Convert.ToDouble(colwidth);
                            rangeDelx.RowHeight = Convert.ToDouble(rowheight);
                        }
                    }
                }
            }
            xlApp.ActiveWindow.DisplayGridlines = false;

            Excel.Range rangeremplace = xlWorkSheet.UsedRange;
            object[,] values8000 = (object[,])rangeremplace.Value2;
            ///////////////row hide "-5"////////////////////////////////////////////////
            for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
            {
                string valuedel = Convert.ToString(values8000[rowhide, col11000]);
                if (Regex.Equals(valuedel, "-5"))
                {
                    Excel.Range rangeDely = xlWorkSheet.Cells[rowhide, col11000] as Excel.Range;
                    rangeDely.EntireRow.Hidden = true;
                }
            }

            Excel.Range rangeDelete = xlWorkSheet.UsedRange.get_Range("Y1", xlWorkSheet.Cells[1, xlWorkSheet.UsedRange.Columns.Count]) as Excel.Range;
            Excel.Range rangeDelete2 = xlWorkSheet.Cells[xlWorkSheet.UsedRange.Rows.Count, 1] as Excel.Range;
            rangeDelete2.EntireRow.Hidden = true;//hide au lieu de supprimer
            rangeDelete.EntireColumn.Hidden = true;


            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            xlWorkBook.Save();
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                textBox20.AppendText("==> Start Formatage des styles de COMPTES ANNUELS dans PrefaceNP : " + System.Environment.NewLine);
               

                int time1 = System.Environment.TickCount;
                if (File.Exists("D:\\ptw\\notepme\\prefaceNP.xlsx"))
                {
                    formartAnuel("D:\\ptw\\notepme\\prefaceNP.xlsx");
                }
                else if (File.Exists("D:\\ptw\\notepme\\prefaceNPs.xlsx"))
                {
                    formartAnuel("D:\\ptw\\notepme\\prefaceNPs.xlsx");
                }
                int time = System.Environment.TickCount;
                time = time - time1;
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Formatage des styles de COMPTES ANNUELS dans PrefaceNP OK : " + timeto + " s");
                MessageBox.Show("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s");
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString()+System.Environment.NewLine);
            }

        }
        // alex the same previews one
        private string button26_Clicktout(object sender, EventArgs e)
        {
            string timeto = "";
            try
            {
                textBox20.AppendText("==> Start Formatage des styles de COMPTES ANNUELS dans PrefaceNP : " + System.Environment.NewLine);
                

                int time1 = System.Environment.TickCount;
                if (File.Exists("D:\\ptw\\notepme\\prefaceNP.xlsx"))
                {
                    formartAnuel("D:\\ptw\\notepme\\prefaceNP.xlsx");
                }
                else if (File.Exists("D:\\ptw\\notepme\\prefaceNPs.xlsx"))
                {
                    formartAnuel("D:\\ptw\\notepme\\prefaceNPs.xlsx");
                }
                int time = System.Environment.TickCount;
                time = time - time1;
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                 timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Formatage des styles de COMPTES ANNUELS dans PrefaceNP OK : " + timeto + " s");
               // MessageBox.Show("Formatage des styles de HISTORIQUE dans PrefaceNP OK : " + timeto + " s");
            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
            return timeto;

        }
        private void checkBox35_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox35.Checked == true)
            {
                checkBox26.Checked = true;
                checkBox27.Checked = true;
                checkBox28.Checked = true;
                checkBox29.Checked = true;
                checkBox30.Checked = true;
                checkBox31.Checked = true;
                checkBox32.Checked = true;
                checkBox33.Checked = true;
                //checkBox34.Checked = true;

            }
            else
            {
                checkBox26.Checked = false;
                checkBox27.Checked = false;
                checkBox28.Checked = false;
                checkBox29.Checked = false;
                checkBox30.Checked = false;
                checkBox31.Checked = false;
                checkBox32.Checked = false;
                checkBox33.Checked = false;
               // checkBox34.Checked = false;
            }
        }

        private void button27_Click_1(object sender, EventArgs e)
        {
            textBox20.Text = "";
        }

        private void button29_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
         
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
           

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            string openfilex = @"D:\ptw\prefaceNP.xlsx";
            xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlWorkSheet);
           int col = cf.FindCodedColumn("9000", range);
           
           FileStream fs = new FileStream(@"D:\ptw\block.index", FileMode.OpenOrCreate);
           StreamWriter sw = new StreamWriter(fs);  
            for (int row = 3; row <= values.GetUpperBound(0); row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "1") || Regex.Equals(value, "-1"))
                {
                    
                  
                    sw.WriteLine(row);
                  

                    
                }
            }
            sw.Close();

            fs.Close();
            xlWorkBook.Close();

            xlApp.Quit();
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (checkBox19.Checked)
            {
                createhissimply(sender, e);
            }
            if (checkBox22.Checked)
            {
                //createanusimply(sender, e);
            }
        }

        //=========================add pagebreak note================================================================================
        private void historypagebreak(object sender, EventArgs e)
        {



            string filePath = @"d:\ptw\index\sheetbreak.index";

            
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
            xlapp.DisplayAlerts = false;
            xlapp.Application.DisplayAlerts = false;
            xlapp.Visible = true;
            Excel.Workbook xlworkbook = xlapp.Workbooks.Open("D:\\ptw\\prefaceNP.xlsx");
            Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");

            Excel.Range xlworksheetused = xlworksheet.UsedRange;
            int col660007000 = 0;
            int row820000 = 0;
            object[,] values = (object[,])xlworksheetused.Value2;
            for (int i = 1; i < xlworksheetused.Columns.Count; i++)
            {
                if (values[xlworksheetused.Rows.Count, i] != null)
                {
                    if (values[xlworksheetused.Rows.Count, i].ToString() == "66000-7000")
                    {
                        col660007000 = i;
                        break;

                    }
                }
            }
            for (int i = 1; i < xlworksheetused.Rows.Count; i++)
            {
                if (values[i, xlworksheetused.Columns.Count] != null)
                {
                    if (values[i, xlworksheetused.Columns.Count].ToString() == "820000")
                    {
                        row820000 = i;
                        break;
                    }
                }
            }

            for (int i = 1; i <= row820000; i++)
            {
                if (values[i, col660007000] != null)
                {
                    if (values[i, col660007000].ToString() == "1")
                    {
                        sw.WriteLine(i);

                    }
                }
            }
            sw.WriteLine(row820000);



            //sw.WriteLine();


            sw.Close();

            file.Close();


            //StreamReader sr = new StreamReader(file);
            //List<string> st=null;
            //while (sr.ReadLine()!=null)
            //{

            //    st.Add(sr.ReadLine());
            //}
        }

        private void historypagebreaks(object sender, EventArgs e)
        {



            string filePath = @"d:\ptw\index\sheetbreaks.index";

            if (File.Exists(filePath))
            {
            }
            else
            {
                File.Create(filePath);
            }
            Thread.Sleep(3000);
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
            xlapp.DisplayAlerts = false;
            xlapp.Application.DisplayAlerts = false;
            xlapp.Visible = true;
            Excel.Workbook xlworkbook = xlapp.Workbooks.Open("D:\\ptw\\prefaceNP.xlsx");
            Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");

            Excel.Range xlworksheetused = xlworksheet.get_Range("A1","BK661");
            int col660007000 = 0;
            int row820000 = 0;
            object[,] values = (object[,])xlworksheetused.Value2;
            for (int i = 1; i < xlworksheetused.Columns.Count; i++)
            {
                if (values[xlworksheetused.Rows.Count, i] != null)
                {
                    if (values[xlworksheetused.Rows.Count, i].ToString() == "10000")
                    {
                        col660007000 = i;
                        break;

                    }
                }
            }
            for (int i = 1; i < xlworksheetused.Rows.Count; i++)
            {
                if (values[i, xlworksheetused.Columns.Count] != null)
                {
                    if (values[i, xlworksheetused.Columns.Count].ToString() == "380000-91000")
                    {
                        row820000 = i;
                        break;
                    }
                }
            }

            for (int i = 1; i <= row820000; i++)
            {
                if (values[i, col660007000] != null)
                {
                    if (values[i, col660007000].ToString() == "1" || values[i, col660007000].ToString() == "-1")
                    {
                        sw.WriteLine(i);

                    }
                }
            }
            sw.WriteLine(row820000);



            //sw.WriteLine();


            sw.Close();

            file.Close();


            //StreamReader sr = new StreamReader(file);
            //List<string> st=null;
            //while (sr.ReadLine()!=null)
            //{

            //    st.Add(sr.ReadLine());
            //}
        }
        private void companuelpagebreak(object sender, EventArgs e)
        {
            string filePath = @"d:\ptw\index\sheetbreakCN.index";
            
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
            xlapp.DisplayAlerts = false;
            xlapp.Application.DisplayAlerts = false;
            xlapp.Visible = true;
            Excel.Workbook xlworkbook = xlapp.Workbooks.Open(@"D:\ptw\prefaceNP.xlsx");
            Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Comptes annuels");

            Excel.Range xlworksheetused = xlworksheet.UsedRange;
            int col32000 = 0;
            int row764000 = 0;

            object[,] values = (object[,])xlworksheetused.Value2;

            for (int i = 1; i <= xlworksheetused.Columns.Count; i++)
            {
                if (values[xlworksheetused.Rows.Count, i] != null)
                {
                    if (values[xlworksheetused.Rows.Count, i].ToString() == "32000")
                    {
                        col32000 = i;
                        break;

                    }
                }
            }
            for (int i = 1; i <= xlworksheetused.Rows.Count; i++)
            {
                if (values[i, xlworksheetused.Columns.Count] != null)
                {
                    if (values[i, xlworksheetused.Columns.Count].ToString() == "764000")
                    {
                        row764000 = i;
                        break;
                    }
                }
            }

            //if(values[1,col8000]!=null)
            //{
            //    if(values[1,col8000].ToString()=="s")
            //    {}
            //}


            for (int i = 1; i <= row764000; i++)
            {
                if (values[i, col32000] != null)
                {
                    if (values[i, col32000].ToString() == "1")
                    {
                        sw.WriteLine(i);

                    }
                }
            }
            sw.WriteLine(row764000);

            sw.Close();

            file.Close();

        }
        private void diviserhis(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;
            fichierprepare = "D:\\ptw\\prefaceNPS.xlsx";
            prefaceNP = "D:\\ptw\\Histo.xlsx";

            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //xlWorkBook = xlApp.Workbooks.Open(fichierprepare, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.UsedRange;
            //petite corr
            object[,] values = (object[,])range.Value2;
            int rCnt = 0;
            int cCnt = 0;
            int row242000 = 0;
          //  CodeFinder cf;
           // cf = new CodeFinder(xlWorkBook, xlWorkSheet);
           // row242000 = cf.FindCodedRow("242000-12000", range);

            //cCnt = range.Columns.Count;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            //{
            //    string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            //    if (Regex.Equals(valuecellabs, "242000-12000"))
            //    {
            //        row242000 = rCnt;
            //        break;
            //    }
            //}
            //Excel.Range cell253F = range.Cells[row242000, 6] as Excel.Range;
            //Excel.Range cell253I = range.Cells[row242000, 9] as Excel.Range;
            //cell253F.Formula = "=C267";
            // cell253I.Formula = "=F267";

            //Hist.Refer coller value
            Excel.Worksheet sheetHistRefer = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer");
            Excel.Range rangeHistRefer = sheetHistRefer.UsedRange;
            rangeHistRefer.Copy(misValue);
            rangeHistRefer.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            ////Hist.Refer mise à zero// dégalage?
            object[,] valuesRefer = (object[,])sheetHistRefer.UsedRange.Value2;

            for (int rowCnt = 1; rowCnt <= rangeHistRefer.Rows.Count - 1; rowCnt++)//sauf derniere ligne..
            {
                string valuecellabs = Convert.ToString(valuesRefer[rowCnt, 1]);
                if (valuecellabs != "")
                {
                    Excel.Range referZero = sheetHistRefer.UsedRange.get_Range(sheetHistRefer.UsedRange.Cells[rowCnt, 3], sheetHistRefer.UsedRange.Cells[rowCnt, 11]) as Excel.Range;
                    //referZero.Copy();
                    referZero.Value2 = 0;
                    //D1 D2 D3  =""
                    if (valuecellabs == "D" || valuecellabs == "D1" || valuecellabs == "d")
                    {
                        referZero.Formula = "=\"\"";
                    }
                }
            }


            //suppression des onglets
            Excel.Worksheet sheetpreface = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface");
            Excel.Worksheet sheetCalculs = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs");
          //  Excel.Worksheet Historiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Worksheet HistPrefacsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Preface-n");
            Excel.Worksheet HistCalculssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Calculs-n");
            Excel.Worksheet HistLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Langues-n");
            Excel.Worksheet HistReferssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Hist.Refer-n");

            Excel.Worksheet ComptesannuelRefssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Annu.Refer");
            Excel.Worksheet Comptesannuelssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
            Excel.Worksheet Osheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("O");
            Excel.Worksheet Identitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Identité");
            Excel.Worksheet Paramimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param impr");
            Excel.Worksheet Psheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("P");
            Excel.Worksheet Paramgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param généraux");
            Excel.Worksheet AdminLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Langues");
            Excel.Worksheet AdminServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Admin.Service");
            Excel.Worksheet Tsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("T");
            Excel.Worksheet ParamSavsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param Sav");
            Excel.Worksheet Macrossheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Macros");
            Excel.Worksheet Vsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("V");
            Excel.Worksheet Mosaiquesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Mosaïque");
            Excel.Worksheet GraphiquesSRsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graphiques SR");
            Excel.Worksheet Graphimprsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Graph impr");
            Excel.Worksheet Dontdeletesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Don't delete");
            Excel.Worksheet Finsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fin");
            Excel.Worksheet ChoixMethodessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixMéthodes");
            Excel.Worksheet Noterecapitulativesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Note récapitulative");
            Excel.Worksheet SyntheseValorisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("SynthèseValorisations");
            Excel.Worksheet DefinitionsArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DéfinitionsArrièrePlan");
            Excel.Worksheet RappelRetraitementssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RappelRetraitements");
            Excel.Worksheet RisqueEntreprisesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("RisqueEntreprise");
            Excel.Worksheet ChoixTauxParamsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ChoixTauxParam");
            Excel.Worksheet TauxParamArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TauxParamArrièrePlan");
            Excel.Worksheet CorrectifsSIGBilansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CorrectifsSIGBilan");
            Excel.Worksheet APNNEsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("APNNE");
            Excel.Worksheet FiscaliteDiffereesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FiscalitéDifférée");
            Excel.Worksheet PatrimonialAncAnccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("PatrimonialAncAncc");
            Excel.Worksheet FondsDeCommercesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("FondsDeCommerce");
            Excel.Worksheet Goodwillsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Goodwill");
            Excel.Worksheet AutresCapitalisationssheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AutresCapitalisations");
            Excel.Worksheet Multiplessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiples");
            Excel.Worksheet MethodesMixtessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("MéthodesMixtes");
            Excel.Worksheet TransactionsComparablessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TransactionsComparables");
            Excel.Worksheet GordonShapiroBatessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("GordonShapiroBates");
            Excel.Worksheet CalculFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CalculFCF");
            Excel.Worksheet DiscountedFCFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("DiscountedFCF");
            Excel.Worksheet CmpcWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWacc");
            Excel.Worksheet CmpcWaccArrierePlansheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CmpcWaccArrièrePlan");
            Excel.Worksheet ModuleWaccsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("ModuleWacc");
            Excel.Worksheet CCEFsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CCEF");
            Excel.Worksheet TriRentabiliteProjetsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TriRentabilitéProjet");
            Excel.Worksheet TourDeTableSynthesesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("TourDeTableSynthèse");
            Excel.Worksheet EvalLanguessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Langues");
            Excel.Worksheet Controlessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles");
            Excel.Worksheet EvalServicesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Eval.Service");
            Excel.Worksheet Composantessheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Composantes");
            Excel.Worksheet Jsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("J");
            Excel.Worksheet Factgenerauxsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Fact généraux");
            Excel.Worksheet Lsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("L");
            Excel.Worksheet Msheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("M");
            Excel.Worksheet Tresoreriesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Trésorerie");
            Excel.Worksheet ABsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("AB");
            Excel.Worksheet Paramtresorsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Param trésor");
            Excel.Worksheet Saisonnalitesheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Saisonnalité");
            Excel.Worksheet Zsheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Z");
            Excel.Worksheet model = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Modèles Goodwill");
          //  Excel.Worksheet delete1 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("PreviNotaPme");
            Excel.Worksheet delete2 = (Excel.Worksheet)xlWorkBook.Sheets.get_Item("Correctifs.Refer");
            Excel.Worksheet sheetCA = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("CA");
            Excel.Worksheet sheetInvestissements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Investissements");
            Excel.Worksheet sheetCpteresultat = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Cpte Résultat");
            Excel.Worksheet sheetFinancements = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Financements");
            Excel.Worksheet sheetbfr = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("BFR");
            Excel.Worksheet sheetbilan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Bilan");
            Excel.Worksheet sheetcontrole2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Contrôles (2)");
            Excel.Worksheet sheetmultiple = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Multiple");
            Excel.Worksheet sheetvalo = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Valo et ouverture du capital");
            Excel.Worksheet sheetplan = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Plan de financement");
            Excel.Worksheet sheetsynthese = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Synthèse SIG et SR");

            sheetCA.Delete();
            sheetInvestissements.Delete();
            sheetCpteresultat.Delete();
            sheetFinancements.Delete();
            sheetbfr.Delete();
            sheetbilan.Delete();
            sheetcontrole2.Delete();
            sheetmultiple.Delete();
            sheetvalo.Delete();
            sheetplan.Delete();
            sheetsynthese.Delete();
          //  delete1.Delete();
            delete2.Delete();
            model.Delete();
            sheetpreface.Delete();
            sheetCalculs.Delete();
            //Historiquesheet.Delete();
            HistPrefacsheet.Delete();
            HistCalculssheet.Delete();
            HistLanguessheet.Delete();
            HistReferssheet.Delete();


            ComptesannuelRefssheet.Delete();
            Comptesannuelssheet.Delete();
            Osheet.Delete();
            Identitesheet.Delete();
            Paramimprsheet.Delete();
            Psheet.Delete();
            Paramgenerauxsheet.Delete();
            AdminLanguessheet.Delete();
            AdminServicesheet.Delete();
            Tsheet.Delete();
            ParamSavsheet.Delete();
            Macrossheet.Delete();
            Vsheet.Delete();
            Mosaiquesheet.Delete();
            GraphiquesSRsheet.Delete();
            Graphimprsheet.Delete();
            Dontdeletesheet.Delete();
            Finsheet.Delete();
            ChoixMethodessheet.Delete();
            Noterecapitulativesheet.Delete();
            SyntheseValorisationssheet.Delete();
            DefinitionsArrierePlansheet.Delete();
            RappelRetraitementssheet.Delete();
            RisqueEntreprisesheet.Delete();
            ChoixTauxParamsheet.Delete();
            TauxParamArrierePlansheet.Delete();
            CorrectifsSIGBilansheet.Delete();
            APNNEsheet.Delete();
            FiscaliteDiffereesheet.Delete();
            PatrimonialAncAnccsheet.Delete();
            FondsDeCommercesheet.Delete();
            Goodwillsheet.Delete();
            AutresCapitalisationssheet.Delete();
            Multiplessheet.Delete();
            MethodesMixtessheet.Delete();
            TransactionsComparablessheet.Delete();
            GordonShapiroBatessheet.Delete();
            CalculFCFsheet.Delete();
            DiscountedFCFsheet.Delete();
            CmpcWaccsheet.Delete();
            CmpcWaccArrierePlansheet.Delete();
            ModuleWaccsheet.Delete();
            CCEFsheet.Delete();
            TriRentabiliteProjetsheet.Delete();
            TourDeTableSynthesesheet.Delete();
            EvalLanguessheet.Delete();
            Controlessheet.Delete();
            EvalServicesheet.Delete();
            Composantessheet.Delete();
            Jsheet.Delete();
            Factgenerauxsheet.Delete();
            Lsheet.Delete();
            Msheet.Delete();
            Tresoreriesheet.Delete();
            ABsheet.Delete();
            Paramtresorsheet.Delete();
            Saisonnalitesheet.Delete();
            Zsheet.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkBook);
            releaseObject(xlApp);


            supprimermoin2Newgai(sender, e);
            //subdiviser 9000
            button4Newgai(sender, e);

            int time2 = System.Environment.TickCount;
            int times = (time2 - time1) / 1000;
            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;

            //timdiviser = Convert.ToString(Convert.ToDecimal(times) / 1000);
            timdiviser = hours + " heures " + minuit + " minutes " + second;

            //MessageBox.Show("jobs done " + tim + " seconds used");

        }
        private void button25_ClickNewgai(object sender, EventArgs e, string s)
        {
            string option = s;
            if (s == ("hisnew"))
            {

                string[] name = { "S-ACT", "S-ANN3", "S-ANN4", "S-ANN5", "S-CR", "S-PAS" };
                //string[] name = { "ANN5", "ANN6", "ANN7", "ANN8", "ANN11", "CR", "PAS" };
                Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
                xlapp.DisplayAlerts = false;
                xlapp.Application.DisplayAlerts = false;
                xlapp.Visible = true;

                //    Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");
                for (int i = 0; i < name.Length; i++)
                {
                    Excel.Workbook xlworkbook = xlapp.Workbooks.Open("D:\\ptw\\notepme\\" + name[i] + ".xlsx");
                    Excel.Workbook xlworkbookstyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");

                    Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");
                    Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookstyle.Worksheets.get_Item("Histo et Histo-s");
                    //xlworksheet.UsedRange.ClearFormats();
                    Excel.Range range = xlworksheet.UsedRange;
                    Excel.Range rangestyle = xlworksheetStyle.UsedRange;

                    int col6000 = 0;
                    int col17000 = 0;
                    //CodeFinder cf = new CodeFinder(xlworkbookstyle, xlworksheetStyle);
                    //col17000 = cf.FindCodedColumn("17000", rangestyle);

                    //cf = new CodeFinder(xlworkbook, xlworksheet);
                    //col6000 = cf.FindCodedColumn("6000", range);

                   

                    object[,] value = (object[,])range.Value2;
                    int col8000=16;
                    for (int t = 1; t <= range.Columns.Count; t++)
                    {
                        if (value[range.Rows.Count, t] != null)
                        {
                            if (value[range.Rows.Count, t].ToString() == "6000")
                            {
                                col6000 = t;
                               
                            }
                             if (value[range.Rows.Count, t].ToString() == "8000")
                            {
                                col8000 = t;
                                break;
                            }

                        }
                    }
                    object[,] valuestyle = (object[,])rangestyle.Value2;
                    for (int t = 1; t <= rangestyle.Columns.Count; t++)
                    {
                        if (valuestyle[rangestyle.Rows.Count, t] != null)
                        {
                            if (valuestyle[rangestyle.Rows.Count, t].ToString() == "17000")
                            {
                                col17000 = t;
                                break;
                            }
                        }

                    }
                    xlworksheetStyle.Cells[8, col17000] = 0;
                   // col6000 = 14;
                    for (int m = 1; m <= range.Rows.Count; m++)
                    {
                        if (value[m, col6000] != null)
                        {
                            if (value[m, col8000] != null && value[m, col8000].ToString() == "-5")
                            {
                            }
                            else
                            {
                                for (int t = 14; t <= 72; t++)
                                {

                                    if (valuestyle[t, col17000] != null)
                                    {


                                        string cherche = xlworksheetStyle.get_Range("A" + t, "A" + t).Value2.ToString();
                                        string fontSize = rangestyle.get_Range(rangestyle.Cells[t, col17000], rangestyle.Cells[t, col17000]).Value2.ToString();

                                        if (cherche == value[m, col6000].ToString())
                                        {
                                            //if(cherche=="8000-150"||cherche=="8000-160"||cherche=="8000-180")
                                            //{
                                            //    string ss="";
                                            //}
                                            Excel.Range changestyle = xlworksheetStyle.get_Range("C" + t, "O" + t);
                                            Excel.Range change = xlworksheet.get_Range("A" + m, "M" + m);

                                            changestyle.Copy();

                                            change.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                            change.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);

                                            change.EntireRow.AutoFit();
                                            change.EntireRow.Font.Size = fontSize;



                                            break;
                                        }
                                    }
                                }
                            }

                        }

                    }
                    xlworkbook.Save();
                    xlworkbook.Close();
                    xlworkbookstyle.Close();
                    releaseObject(xlworkbook);
                    releaseObject(xlworkbookstyle);
                    releaseObject(xlworksheet);
                    releaseObject(xlworksheetStyle);

                }
                xlapp.Quit();
            }
            else if (s == "ann")
            {
                string[] name = { "ANNUEL-CR.xlsx", "ANNUEL-BILAN-FR.xlsx", "ANNUEL-BILAN-ENG.xlsx", "ANNUEL-FR-BFR-TRES.xlsx", "ANNUEL-FLUX-TRES.xlsx", "ANNUEL-RATIOS.xlsx", "ANNUEL-SYNTHESIS.xlsx" };

                //string[] name = { "ANNUEL-FR-BFR-TRES.xlsx"};
                Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
                xlapp.DisplayAlerts = false;
                xlapp.Application.DisplayAlerts = false;
                xlapp.Visible = true;

                //    Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");
                for (int i = 0; i < name.Length; i++)
                {
                    Excel.Workbook xlworkbook = xlapp.Workbooks.Open("D:\\ptw\\notepme\\" + name[i]);
                    Excel.Workbook xlworkbookstyle = xlapp.Workbooks.Open("D:\\ptw\\style nota-pme.xlsx");

                    Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Comptes annuels");
                    Excel.Worksheet xlworksheetStyle = (Excel.Worksheet)xlworkbookstyle.Worksheets.get_Item("Annuel");
                    Excel.Range range = xlworksheet.UsedRange;
                    Excel.Range rangestyle = xlworksheetStyle.UsedRange;

                    int col90001000 = 0;
                    int col27000 = 0;
                    //CodeFinder cf = new CodeFinder(xlworkbookstyle, xlworksheetStyle);
                    //col27000 = cf.FindCodedColumn("27000", rangestyle);

                    //cf = new CodeFinder(xlworkbook, xlworksheet);
                    //col90001000 = cf.FindCodedColumn("9000-1000", range);

                    //xlworksheetStyle.Cells[2, col27000] = 0;
                    object[,] value = (object[,])range.Value2;
                    for (int t = 1; t <= range.Columns.Count; t++)
                    {
                        if (value[range.Rows.Count, t] != null)
                        {
                            if (value[range.Rows.Count, t].ToString() == "9000-1000")
                            {
                                col90001000 = t;
                                break;
                            }
                        }
                    }
                    object[,] valuestyle = (object[,])rangestyle.Value2;
                    for (int t = 1; t <= rangestyle.Columns.Count; t++)
                    {
                        if (valuestyle[rangestyle.Rows.Count, t] != null)
                        {
                            if (valuestyle[rangestyle.Rows.Count, t].ToString() == "27000")
                            {
                                col27000 = t;
                                break;
                            }
                        }

                    }
                    for (int m = 1; m <= range.Rows.Count; m++)
                    {
                        if (value[m, col90001000] != null)
                        {
                            for (int t = 7; t <= 45; t++)
                            {

                                if (valuestyle[t, col27000] != null)
                                {


                                    string cherche = xlworksheetStyle.get_Range("A" + t, "A" + t).Value2.ToString();
                                    //string fontSize = rangestyle.get_Range(rangestyle.Cells[t, col27000], rangestyle.Cells[t, col27000]).Value2.ToString();

                                    if (cherche == range.get_Range(range.Cells[m, col90001000], range.Cells[m, col90001000]).Value2.ToString())
                                    {
                                        Excel.Range rangestylee = xlworksheetStyle.get_Range("C" + t, "Z" + t);
                                        Excel.Range change = xlworksheet.get_Range("A" + m, "X" + m);

                                        rangestylee.Copy();
                                        change.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                        change.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                                    }
                                }
                            }
                        }

                    }
                    xlworkbook.Save();
                    xlworkbook.Close();
                    xlworkbookstyle.Close();
                    releaseObject(xlworkbook);
                    releaseObject(xlworkbookstyle);
                    releaseObject(xlworksheet);
                    releaseObject(xlworksheetStyle);

                }
                xlapp.Quit();
            }

        }
        private void button19_ClickNewgai(object sender, EventArgs e, string s)
        {//pas3 not worked
            string option = s;
            pathstylerfinal = "D:\\ptw\\notepme";
            if (s == "hisnew")
            {
                string[] namestable = { "S-ACT", "S-PAS", "S-CR", "S-ANN3", "S-ANN4", "S-ANN5", };
                //string[] namestable = { "ANN5", "ANN6", "ANN7", "ANN8", "ANN11", "CR", "PAS" };
                int rcont = 1;
                int ccont = 12;

                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";

                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;
                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
                    //model.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Historique");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Historique");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Historique");


                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
                    //set using language to be english
                    Excel.Range rangeEng = enlanguesheet.UsedRange;
                    object[,] valueEng = (object[,])rangeEng.Value2;
                    int row12950005000e = 0;
                    for (int t = 1; t <= rangeEng.Rows.Count; t++)
                    {
                        if (valueEng[t, rangeEng.Columns.Count] != null)
                        {
                            if (valueEng[t, rangeEng.Columns.Count].ToString() == "1291000")
                            {
                                row12950005000e = t;
                                break;
                            }
                        }
                    }


                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E" + row12950005000e);
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B" + row12950005000e);
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range rangeGer = glanguesheet.UsedRange;
                    object[,] valueGer = (object[,])rangeGer.Value2;
                    int row12950005000g = 0;
                    for (int t = 1; t <= rangeGer.Rows.Count; t++)
                    {
                        if (valueGer[t, rangeGer.Columns.Count] != null)
                        {
                            if (valueGer[t, rangeGer.Columns.Count].ToString() == "1291000")
                            {
                                row12950005000g = t;
                                break;
                            }
                        }
                    }

                    Excel.Range grange = glanguesheet.get_Range("F4", "f" + row12950005000g);
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B" + row12950005000g);
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //    string strcell = values[rcont, ccont].ToString();

                    try
                    {
                        //fr 
                        copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //en
                        copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //german
                        copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");

                    deletelaguage.Delete();

                    enlanguesheet.Delete();
                    glanguesheet.Delete();


                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();

                    releaseObject(deleteworksheet1);
                    releaseObject(deleteworksheet2);
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);

                    app.Quit();

                }
            }
            if (option == "ann")
            {
                string[] namestable = { "ANNUEL-CR.xlsx", "ANNUEL-BILAN-FR.xlsx", "ANNUEL-BILAN-ENG.xlsx", "ANNUEL-FR-BFR-TRES.xlsx", "ANNUEL-FLUX-TRES.xlsx", "ANNUEL-RATIOS.xlsx", "ANNUEL-SYNTHESIS.xlsx" };
                //string[] namestable = {  "ANNUEL-FR-BFR-TRES"};
                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";
                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;

                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Comptes annuels");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Comptes annuels");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Comptes annuels");

                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Admin.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Admin.Langues");
                    //set using language to be english
                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E3511");
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B3511");
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range grange = glanguesheet.get_Range("F4", "F3511");
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B3511");
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    int rcont = 1;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //string strcell = values[rcont, ccont].ToString();

                    try
                    {
                        //fr 
                        copyrange = myworksheet.get_Range(myworksheet.Cells[rcont, 1], myworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //en
                        copyrange = enworksheet.get_Range(enworksheet.Cells[rcont, 1], enworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //german
                        copyrange = gworksheet.get_Range(gworksheet.Cells[rcont, 1], gworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Admin.Langues");
                    deletelaguage.Delete();
                    enlanguesheet.Delete();
                    glanguesheet.Delete();
                    Excel.Worksheet delosheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet2 = (Excel.Worksheet)enworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet3 = (Excel.Worksheet)gworkbook.Worksheets.get_Item("O");
                    delosheet1.Delete();
                    delosheet2.Delete();
                    delosheet3.Delete();
                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    deletesheets(myworkbook);
                    deletesheets(enworkbook);
                    deletesheets(gworkbook);
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);


                    app.Quit();

                }

            }

            //else if (checkBox23.Checked)
            //{

            //    string[] namestable = { "EVAL-SYNTHVALO2", "EVAL-SYNTHVALO1", "EVAL-SYNTHMULT1" };
            //    int rcont = 1;
            //    int ccont = 12;

            //    object misValue = System.Reflection.Missing.Value;
            //    for (int i = 0; i < namestable.Count(); i++)
            //    {
            //        string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
            //        string enpath = "d:\\ptw\\notepme\\" + namestable[i] + "_EN.xlsx";
            //        string gpath = "d:\\ptw\\notepme\\" + namestable[i] + "_GER.xlsx";
            //        string frpath = "d:\\ptw\\notepme\\" + namestable[i] + "_FR.xlsx";

            //        Excel.Application app = new Excel.Application();
            //        app.DisplayAlerts = false;
            //        app.Visible = true;
            //        Excel.Workbook myworkbook;
            //        Excel.Workbook enworkbook;
            //        Excel.Workbook gworkbook;
            //        Excel.Worksheet myworksheet;
            //        Excel.Worksheet enworksheet;
            //        Excel.Worksheet gworksheet;
            //        Excel._Worksheet deleteworksheet1;
            //        Excel._Worksheet deleteworksheet2;

            //        myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        // deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
            //        //  deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
            //        // deleteworksheet1.Delete();
            //        //deleteworksheet2.Delete();
            //        //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
            //        //model.Delete();
            //        myworkbook.SaveCopyAs(frpath);
            //        myworkbook.SaveCopyAs(enpath);
            //        myworkbook.SaveCopyAs(gpath);
            //        myworkbook.Close();
            //        myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //        myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("SynthèseValorisations");
            //        enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("SynthèseValorisations");
            //        gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("SynthèseValorisations");


            //        //Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
            //        //Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
            //        ////set using language to be english
            //        //Excel.Range enrange = enlanguesheet.get_Range("E4", "E1043");
            //        //Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B1043");
            //        //enrange.Copy(enpasterange);
            //        //releaseObject(enrange);
            //        //releaseObject(enpasterange);
            //        ////set using language to be german
            //        //Excel.Range grange = glanguesheet.get_Range("F4", "f1043");
            //        //Excel.Range gpasterange = glanguesheet.get_Range("B4", "B1043");
            //        //grange.Copy(gpasterange);
            //        //releaseObject(grange);
            //        //releaseObject(gpasterange);

            //        //Excel.Range userange = myworksheet.UsedRange;
            //        //object[,] values = (object[,])userange.Value2;
            //        //Excel.Range copyrange;
            //        ////for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
            //        ////{
            //        ////    string strcell = values[rcont, ccont].ToString();

            //        //try
            //        //{
            //        //    //fr 
            //        //    copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //        //    //en
            //        //    copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //        //    //german
            //        //    copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


            //        //}
            //        //catch (Exception ex)
            //        //{
            //        //}
            //        //}
            //        //}

            //        //copyrange.Copy(copyrange);
            //        //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //deletesheet.Delete();
            //        //deleteensheet.Delete();
            //        //deletegsheet.Delete();
            //        //Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");
            //        // deletelaguage.Delete();

            //        //enlanguesheet.Delete();
            //        // glanguesheet.Delete();


            //        myworkbook.Save();
            //        enworkbook.Save();
            //        gworkbook.Save();
            //        myworkbook.Close();
            //        enworkbook.Close();
            //        gworkbook.Close();

            //        //releaseObject(deleteworksheet1);
            //        // releaseObject(deleteworksheet2);
            //        // releaseObject(deleteensheet);
            //        // releaseObject(enlanguesheet);
            //        releaseObject(enworksheet);
            //        releaseObject(enworkbook);
            //        // releaseObject(deletegsheet);
            //        // releaseObject(glanguesheet);
            //        releaseObject(gworksheet);
            //        releaseObject(gworkbook);
            //        // releaseObject(deletesheet);
            //        //releaseObject(deletelaguage);
            //        releaseObject(myworksheet);
            //        releaseObject(myworkbook);

            //        app.Quit();

            //    }

        }
        private void hiddengai(string flag)
        {
            string path = "D:\\ptw\\notepme\\";
            string[] filename = Directory.GetFiles(path);
            Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
            xlapp.DisplayAlerts = false;
            xlapp.Application.DisplayAlerts = false;
            xlapp.Visible = true;

            for (int i = 0; i < filename.Length; i++)
            {
                string name = filename[i];
                if (!name.Contains("ANNUEL") && !name.Contains("6") && !name.Contains("7") && !name.Contains("8") && !name.Contains("11") && name.Contains("S-"))
                {
                    if (flag == "hisnew")
                    {
                        Excel.Workbook xlworkbook = xlapp.Workbooks.Open(name);
                        Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");
                        Excel.Range hide1 = xlworksheet.get_Range("N1", xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, xlworksheet.UsedRange.Columns.Count]);
                        hide1.EntireColumn.Hidden = true;

                        Excel.Range hide2 = xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1] as Excel.Range;
                        hide2.EntireRow.Hidden = true;

                        //Excel.Range range1 = xlworksheet.get_Range("A1");
                        //double height = double.Parse(range1.Height.ToString());

                        //for (int n = 1; n <= 4; n++)
                        //{
                        //    Excel.Range rangex = xlworksheet.get_Range("A" + n).EntireRow;
                        //    rangex.EntireRow.RowHeight = height /4;
                        //}


                        xlapp.ActiveWindow.DisplayGridlines = false;
                        xlworkbook.Save();
                        xlworkbook.Close();

                        releaseObject(xlworkbook);

                        releaseObject(xlworksheet);
                    }
                }
                else
                {
                    if (flag == "ann")
                    {
                        Excel.Workbook xlworkbook = xlapp.Workbooks.Open(name);
                        Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Comptes annuels");
                        Excel.Range hide1 = xlworksheet.get_Range("Y1", xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, xlworksheet.UsedRange.Columns.Count]);
                        hide1.EntireColumn.Hidden = true;

                        Excel.Range hide2 = xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1] as Excel.Range;
                        hide2.EntireRow.Hidden = true;
                        xlapp.ActiveWindow.DisplayGridlines = false;
                        xlworkbook.Save();
                        xlworkbook.Close();

                        releaseObject(xlworkbook);

                        releaseObject(xlworksheet);
                    }
                }
            }
            xlapp.Quit();
        }
        private void supprimermoin2Newgai(object sender, EventArgs e)
        {
            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBook = xlApp.Workbooks.Open(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.get_Range("A1","W661");
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////944000//////////////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row944000 = range.Rows.Count;
            //int toalrows = 654;
            //for (int i = 1;i<=300 ;i++ )
            //{
            //    string valuecellabs = Convert.ToString(values[toalrows, i]);
            //    if (Regex.Equals(valuecellabs, "400000"))
            //    {
            //        row944000 = rCnt;
            //        break;
            //    }
            //}
            cCnt = range.Columns.Count;
            //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            // {
            // string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
            // if (Regex.Equals(valuecellabs, "400000"))
            // {
            //  row944000 = rCnt;
            //  break;
            // }
            // }

            //for (int col = 1; col <= xlWorkSheet.UsedRange.Columns.Count; col++)
            //{
            //    string value = Convert.ToString(values[row944000, col]);
            //    if (Regex.Equals(value, "-2"))
            //    {
            //        Excel.Range rangeDelx = xlWorkSheet.Cells[row944000, col] as Excel.Range;

            //        rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

            //        range = xlWorkSheet.UsedRange;
            //        values = (object[,])range.Value2;
            //        col--;
            //    }
            //}
            int nubmer = 0;

            for (int col = 1; col <= range.Columns.Count; col++)
            {

                string value = Convert.ToString(values[row944000-1, col]);

                if (Regex.Equals(value, "-2"))
                {
                    nubmer++;
                }
                else
                {
                    if (nubmer != 0)
                    {
                        Excel.Range rangeDelx = xlWorkSheet.get_Range(xlWorkSheet.Cells[row944000-1, col - nubmer], xlWorkSheet.Cells[row944000, col - 1]) as Excel.Range;

                        rangeDelx.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                        range = xlWorkSheet.UsedRange;
                        values = (object[,])range.Value2;
                        col = col - nubmer;
                    }
                    nubmer = 0;
                }

            }
            range = xlWorkSheet.UsedRange;
            cCnt = range.Columns.Count;
            values = (object[,])range.Value2;
            for (int col = 1; col <= cCnt; col++)
            {
                if (values[row944000, col] != null)
                {
                    string value = Convert.ToString(values[row944000, col]);
                    if (Regex.Equals(value, "-4"))
                    {
                        Excel.Range rangeEffacer = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col], xlWorkSheet.Cells[row944000 - 1, col]) as Excel.Range;
                        rangeEffacer.ClearContents();
                    }

                }
            }



            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlApp.DisplayAlerts = true;
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string tim = Convert.ToString(Convert.ToDecimal(times) / 1000);
            //MessageBox.Show("jobs done " + tim + " seconds used");

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void button4Newgai(object sender, EventArgs e)
        {
            pathnotapme = "D:\\ptw\\notepme";
            //pathstylerfinal = "D:\\ptw\\changeStyle\\divi\\final";

            string openfilex = "D:\\ptw\\Histo.xlsx";

            ////////////////open excel///////////////////////////////////////
            Thread.Sleep(3000);
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookx1;
            Excel.Workbook xlWorkBooknewx1;
            object misValue = System.Reflection.Missing.Value;
            //////////creat modele histox.xls pour fichier diviser////////////////////////////////
            Excel.Application xlAppRef;
            Excel.Workbook xlWorkBookRef;
            xlAppRef = new Excel.ApplicationClass();
            xlAppRef.Visible = true;
            xlAppRef.DisplayAlerts = false;
            xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlWorkBookRef = xlAppRef.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);


            Excel.Worksheet xlWorkSheetRef = (Excel.Worksheet)xlWorkBookRef.Worksheets.get_Item("Historique");
            Excel.Range rangeRefall = xlWorkSheetRef.get_Range("A1","X661");
            object[,] valuess =(object[,])rangeRefall.Value2;
            //bug : le seul moyen pour supprimer la dernière colonne est de chnager la largeur de toutes les colonnes (on ne sait pas pourquoi) !!!
          
            xlWorkSheetRef.Cells.ColumnWidth = 20;
            int rowcount = rangeRefall.Rows.Count;
            int colcount = rangeRefall.Columns.Count;
           
            Excel.Range rangeRef = xlWorkSheetRef.get_Range("A661", "X661");
            rangeRef.EntireRow.Copy(misValue);
           
            rangeRef.EntireRow.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
            Excel.Range rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("X1", "BV1") as Excel.Range;
            rangeRefdel.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A662", "X" + xlWorkSheetRef.UsedRange.Rows.Count) as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
           
            rangeRefdel = xlWorkSheetRef.UsedRange.get_Range("A1", "X660") as Excel.Range;
            rangeRefdel.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            Excel.Range rangeA1 = xlWorkSheetRef.Cells[1, 1] as Excel.Range;
            rangeA1.Activate();
            
            xlWorkSheetRef.SaveAs("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBookRef.Close(true, misValue, misValue);
            xlAppRef.Quit();
            //////////////////////////////////////////////////////////////////////////////////
            Thread.Sleep(3000);
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlApp.Application.DisplayAlerts = false;

            //MessageBox.Show(openfilex);//D:\ptw\Histo.xls
            string remplacehisto8 = "[" + openfilex.Substring(7, 9) + "]";
            //xlWorkBook = xlApp.Workbooks.Open(openfilex, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            xlWorkBook = xlApp.Workbooks.Open(openfilex, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            Excel.Range range = xlWorkSheet.get_Range("A1", "W661");
            object[,] values = (object[,])range.Value2;



            int rCnt = 0;
            int rowx = xlWorkSheet.get_Range("A1", "W661").Rows.Count; 
            int cCnt = 0;
            int col = 0;
            int col3000 = 0;
            int col4000 = 0;
            int col5000 = 0;
            int col8000 = 0;
            int col83000 = 0;
            rCnt = xlWorkSheet.get_Range("A1", "W661").Rows.Count; 

           // CodeFinder cf;
           // cf = new CodeFinder(xlWorkBook, xlWorkSheet);
           // col3000 = cf.FindCodedColumn("3000", range);
           // col4000 = cf.FindCodedColumn("4000", range);
          //  col5000 = cf.FindCodedColumn("5000", range);
          //  col8000 = cf.FindCodedColumn("8000", range);
          //  col = cf.FindCodedColumn("10000", range);
           // col83000 = cf.FindCodedColumn("83000", range);
            //if (values[1, col].ToString() != null)
            //{
            //    if (values[1, col].ToString() == "S")
            //    {
            //        xlWorkSheet.Cells[1, col] = 1;

            //        //xlWorkSheet.get_Range(xlWorkSheet.Cells[1, col]).Value2 ="l" ;
            //    }
            //}

            for (cCnt = 1; cCnt <= xlWorkSheet.get_Range("A1", "W661").Columns.Count - 1 ; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "3000"))
                {
                    col3000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "4000"))
                {
                    col4000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "5000"))
                {
                    col5000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "8000"))
                {
                    col8000 = cCnt;
                }
                if (Regex.Equals(valuecellabs, "10000"))
                {
                    col = cCnt;
                }
                if (Regex.Equals(valuecellabs, "83000"))
                {
                    col83000 = cCnt;
                    break;
                }
            }
            int fileflag = 0;
            for (int row = 25; row <= range.Rows.Count - 1; row++)
            {
                string value = Convert.ToString(values[row, col]);
                if (Regex.Equals(value, "-1"))
                {
                    Thread.Sleep(3000);
                    xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    // xlWorkBookx1 = xlApp.Workbooks.Open("D:\\ptw\\Histox.xlsx", misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    Excel.Worksheet xlWorkSheetx1 = (Excel.Worksheet)xlWorkBookx1.Worksheets.get_Item("Historique");
                    string[] namestable = { "S-ACT.xlsx", "S-PAS.xlsx", "S-CR.xlsx", "S-ANN3.xlsx", "S-ANN4.xlsx", "S-ANN5.xlsx" };

                    string divisavenom = pathnotapme + "\\" + namestable[fileflag];
                    divitylerfinal = pathstylerfinal + "\\" + namestable[fileflag];
                    System.IO.Directory.CreateDirectory(pathnotapme);//////////////cree repertoire

                    xlWorkSheetx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    xlWorkBookx1.Close(true, misValue, misValue);
                    ////////////Grande titre "-1"/////////////////////////////////////////////////////////////////
                    if (Regex.Equals(Convert.ToString(values[25, col]), "-1"))
                    {
                        Excel.Range rangegtitre = xlWorkSheet.Cells[25, col] as Excel.Range;
                        Excel.Range rangePastegtitre = xlWorkSheet.UsedRange.Cells[24, 1] as Excel.Range;
                        rangegtitre.EntireRow.Cut(rangePastegtitre.EntireRow);

                        Excel.Range rangegtitreblank = xlWorkSheet.Cells[25, col] as Excel.Range;
                        rangegtitreblank.EntireRow.Delete(misValue);
                        row--;// point important, pour garder l'ordre de ligne ne change pas
                    }

                    ////////////////////insertion///////////////////////////////////////////////////////////////////
                    Excel.Range rangeDelx = xlWorkSheet.Cells[row, col] as Excel.Range;
                    Excel.Range rangediviser = xlWorkSheet.UsedRange.get_Range("A1", xlWorkSheet.Cells[row - 1, col-1]) as Excel.Range;
                    Excel.Range rangedelete = xlWorkSheet.UsedRange.get_Range("A25", xlWorkSheet.Cells[row - 1, col]) as Excel.Range;
                    rangediviser.EntireRow.Select();
                    rangediviser.EntireRow.Copy(misValue);
                    //MessageBox.Show(row.ToString());

                    xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //xlWorkBooknewx1 = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;


                    Excel.Worksheet xlWorkSheetnewx1 = (Excel.Worksheet)xlWorkBooknewx1.Worksheets.get_Item("Historique");
                    //xlWorkBooknewx1.set_Colors(misValue, xlWorkBook.get_Colors(misValue));
                    Excel.Range rangenewx1 = xlWorkSheetnewx1.Cells[1, 1] as Excel.Range;
                    rangenewx1.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, misValue);

                    xlWorkSheetnewx1.SaveAs(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);

                    //modifier lien pour effacer cross file reference!!!!!!!!!!!!!!2003-2010
                    xlWorkBooknewx1.ChangeLink(openfilex, divisavenom);
                    xlWorkBooknewx1.Close(true, misValue, misValue);

                    ////////////////////replace formulaire contient ptw/histo8.xls///////////////////
                    Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    //Excel.Workbook xlWorkBookremplace = xlApp.Workbooks.Open(divisavenom, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
                    xlApp.DisplayAlerts = false;
                    xlApp.Application.DisplayAlerts = false;



                    Excel.Worksheet xlWorkSheetremplace = (Excel.Worksheet)xlWorkBookremplace.Worksheets.get_Item("Historique");
                    Excel.Range rangeremplace = xlWorkSheetremplace.get_Range("A1", "W661");
                    rangeremplace.Cells.Replace(remplacehisto8, "", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);//NB remplacehisto8 il faut ameliorer pour adapder tous les cas
                    ////////delete col8000 "-2"//////////////////////////////////////////////////
                    object[,] values8000 = (object[,])rangeremplace.Value2;

                    for (int rowdel = 1; rowdel <= rangeremplace.Rows.Count; rowdel++)
                    {
                        string valuedel = Convert.ToString(values8000[rowdel, col8000]);
                        if (Regex.Equals(valuedel, "-2"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowdel, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowdel--;
                        }
                    }
                    ///////////////row hide "-5"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-5"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Hidden = true;
                        }
                    }
                    ///////////////row supprimer "-6"////////////////////////////////////////////////
                    for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    {
                        string valuedel = Convert.ToString(values8000[rowhide, col8000]);
                        if (Regex.Equals(valuedel, "-6"))
                        {
                            Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col8000] as Excel.Range;
                            rangeDely.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                            rangeremplace = xlWorkSheetremplace.UsedRange;
                            values8000 = (object[,])rangeremplace.Value2;
                            rowhide--;
                        }
                    }
                    ///////////////Hide -1 pour col 83000/////////////////////////////////////////////
                    //for (int rowhide = 1; rowhide <= rangeremplace.Rows.Count; rowhide++)
                    //{
                    //    string valuedel = Convert.ToString(values8000[rowhide, col83000]);
                    //    if (Regex.Equals(valuedel, "-1"))
                    //    {
                    //        Excel.Range rangeDely = xlWorkSheetremplace.Cells[rowhide, col83000] as Excel.Range;
                    //        rangeDely.EntireRow.Hidden = true;
                    //    }
                    //}
                    /////////////////////////////////////////////////////////////////////////////////
                    object[,] valuesNX = (object[,])rangeremplace.Value2;
                    //string valueNX = Convert.ToString(valuesNX[row, col]);
                    for (int row3000 = 1; row3000 <= rangeremplace.Rows.Count; row3000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row3000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row3000, col3000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row3000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row4000 = 1; row4000 <= rangeremplace.Rows.Count; row4000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row4000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row4000, col4000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row4000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }
                    for (int row5000 = 1; row5000 <= rangeremplace.Rows.Count; row5000++)
                    {
                        Excel.Range rangeprey = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-3"))
                        {
                            rangeprey.Locked = false;
                            rangeprey.FormulaHidden = false;
                        }
                        if (Regex.Equals(Convert.ToString(valuesNX[row5000, col8000]), "-4"))
                        {
                            rangeprey.Value2 = 0;
                            rangeprey.Locked = true;
                            rangeprey.FormulaHidden = true;
                        }
                        Excel.Range rangeDely = xlWorkSheetremplace.Cells[row5000, col5000] as Excel.Range;
                        if (rangeDely.Locked.ToString() != "True" && Convert.ToString(valuesNX[row5000, col8000]) != "-7")//-7 non zero
                        {
                            rangeDely.Value2 = 0;
                        }
                    }

                    ////////////////////////////////////////////////////////////////////////////
                    xlApp.ActiveWindow.SplitRow = 0;
                    xlApp.ActiveWindow.SplitColumn = 0;
                    xlWorkBookremplace.Save();
                    xlWorkBookremplace.Close(true, misValue, misValue);
                    //if (checkBox20.Checked == true)
                    //{
                    //    fileAstyler = divisavenom;
                    //    Xmllire_Click(sender, e);
                    //}

                    rangedelete.Copy(misValue);
                    rangedelete.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    range = xlWorkSheet.UsedRange;
                    values = (object[,])range.Value2;
                    row = 25;//important remise le ligne commencer apres action delete 1:)25ligne
                    xlWorkSheet.Activate();
                    fileflag++;
                }
            }
            xlApp.Quit();

            //MessageBox.Show("jobs done");
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }
        private void createhissimply(object sender, EventArgs e)
        {
            string his = "hisnew";
            diviserhis(sender, e);
            button25_ClickNewgai(sender, e, his);

            button19_ClickNewgai(sender, e, his);
            hiddengai(his);

        }
        private void createanusimply(object sender, EventArgs e)
        {
            string ann = "ann";
            //button25_ClickNewgai(sender, e, ann);

            //button25_ClickNew(sender, e, ann);
            button19_ClickNew(sender, e, ann);
            hidden(ann);
        }
        private void button19_ClickNew(object sender, EventArgs e, string s)
        {//pas3 not worked
            string option = s;
            pathstylerfinal = "D:\\ptw\\notepme";
            if (s == "his")
            {
                string[] namestable = { "ACT", "PAS", "CR", "ANN5", "ANN6", "ANN7", "ANN8", "ANN11" };
                //string[] namestable = { "ANN5", "ANN6", "ANN7", "ANN8", "ANN11", "CR", "PAS" };
                int rcont = 1;
                int ccont = 12;

                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";

                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;
                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
                    //model.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Historique");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Historique");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Historique");


                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
                    //set using language to be english
                    Excel.Range rangeEng = enlanguesheet.UsedRange;
                    object[,] valueEng = (object[,])rangeEng.Value2;
                    int row12950005000e = 0;
                    for (int t = 1; t <= rangeEng.Rows.Count; t++)
                    {
                        if (valueEng[t, rangeEng.Columns.Count] != null)
                        {
                            if (valueEng[t, rangeEng.Columns.Count].ToString() == "1295000-5000")
                            {
                                row12950005000e = t;
                                break;
                            }
                        }
                    }


                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E" + row12950005000e);
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B" + row12950005000e);
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range rangeGer = glanguesheet.UsedRange;
                    object[,] valueGer = (object[,])rangeGer.Value2;
                    int row12950005000g = 0;
                    for (int t = 1; t <= rangeGer.Rows.Count; t++)
                    {
                        if (valueGer[t, rangeGer.Columns.Count] != null)
                        {
                            if (valueGer[t, rangeGer.Columns.Count].ToString() == "1295000-5000")
                            {
                                row12950005000g = t;
                                break;
                            }
                        }
                    }

                    Excel.Range grange = glanguesheet.get_Range("F4", "f" + row12950005000g);
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B" + row12950005000g);
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //    string strcell = values[rcont, ccont].ToString();

                    try
                    {
                        //fr 
                        copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //en
                        copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //german
                        copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[userange.Rows.Count, 2]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");

                    deletelaguage.Delete();

                    enlanguesheet.Delete();
                    glanguesheet.Delete();


                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();

                    releaseObject(deleteworksheet1);
                    releaseObject(deleteworksheet2);
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);

                    app.Quit();

                }
            }
            if (option == "ann")
            {
                string[] namestable = { "ANNUEL-CR", "ANNUEL-BILAN-FR", "ANNUEL-BILAN-ENG", "ANNUEL-FR-BFR-TRES", "ANNUEL-FLUX-TRES", "ANNUEL-RATIOS", "ANNUEL-SYNTHESIS" };
                //string[] namestable = {  "ANNUEL-FR-BFR-TRES"};
                object misValue = System.Reflection.Missing.Value;
                for (int i = 0; i < namestable.Count(); i++)
                {
                    string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
                    string enpath = pathstylerfinal + "\\" + namestable[i] + "_EN.xlsx";
                    string gpath = pathstylerfinal + "\\" + namestable[i] + "_GER.xlsx";
                    string frpath = pathstylerfinal + "\\" + namestable[i] + "_FR.xlsx";
                    Excel.Application app = new Excel.Application();
                    app.DisplayAlerts = false;
                    app.Visible = true;

                    Excel.Workbook myworkbook;
                    Excel.Workbook enworkbook;
                    Excel.Workbook gworkbook;
                    Excel.Worksheet myworksheet;
                    Excel.Worksheet enworksheet;
                    Excel.Worksheet gworksheet;
                    Excel._Worksheet deleteworksheet1;
                    Excel._Worksheet deleteworksheet2;

                    myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
                    deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
                    deleteworksheet1.Delete();
                    deleteworksheet2.Delete();
                    myworkbook.SaveCopyAs(frpath);
                    myworkbook.SaveCopyAs(enpath);
                    myworkbook.SaveCopyAs(gpath);
                    myworkbook.Close();
                    myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                    myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Comptes annuels");
                    enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Comptes annuels");
                    gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Comptes annuels");

                    Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Admin.Langues");
                    Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Admin.Langues");
                    //set using language to be english
                    Excel.Range enrange = enlanguesheet.get_Range("E4", "E3511");
                    Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B3511");
                    enrange.Copy(enpasterange);
                    releaseObject(enrange);
                    releaseObject(enpasterange);
                    //set using language to be german
                    Excel.Range grange = glanguesheet.get_Range("F4", "F3511");
                    Excel.Range gpasterange = glanguesheet.get_Range("B4", "B3511");
                    grange.Copy(gpasterange);
                    releaseObject(grange);
                    releaseObject(gpasterange);

                    Excel.Range userange = myworksheet.UsedRange;
                    object[,] values = (object[,])userange.Value2;
                    Excel.Range copyrange;
                    int rcont = 1;
                    //for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
                    //{
                    //string strcell = values[rcont, ccont].ToString();

                    try
                    {
                        //fr 
                        copyrange = myworksheet.get_Range(myworksheet.Cells[rcont, 1], myworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //en
                        copyrange = enworksheet.get_Range(enworksheet.Cells[rcont, 1], enworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                        //german
                        copyrange = gworksheet.get_Range(gworksheet.Cells[rcont, 1], gworksheet.Cells[userange.Rows.Count, 1]);
                        copyrange.Copy(misValue);
                        copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


                    }
                    catch (Exception ex)
                    {
                    }
                    //}
                    //}

                    //copyrange.Copy(copyrange);
                    //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
                    //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
                    //deletesheet.Delete();
                    //deleteensheet.Delete();
                    //deletegsheet.Delete();
                    Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Admin.Langues");
                    deletelaguage.Delete();
                    enlanguesheet.Delete();
                    glanguesheet.Delete();
                    Excel.Worksheet delosheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet2 = (Excel.Worksheet)enworkbook.Worksheets.get_Item("O");
                    Excel.Worksheet delosheet3 = (Excel.Worksheet)gworkbook.Worksheets.get_Item("O");
                    delosheet1.Delete();
                    delosheet2.Delete();
                    delosheet3.Delete();

                    //for (int k = 1; k <= myworkbook.Worksheets.Count;k++ )
                    //{
                    //    Excel.Worksheet w =(Excel.Worksheet) myworkbook.Worksheets.get_Item(k);
                    //    if (w.Name.ToString() != "Comptes annuels" && w.Name.ToString() != "Annu.Refer")
                    //    {
                    //        w.Delete();
                    //    }
                    //}
                    //for (int k = 1; k <= enworkbook.Worksheets.Count; k++)
                    //{
                    //    Excel.Worksheet w = (Excel.Worksheet)enworkbook.Worksheets.get_Item(k);
                    //    if (w.Name.ToString() != "Comptes annuels" && w.Name.ToString() != "Annu.Refer")
                    //    {
                    //        w.Delete();
                    //    }
                    //}
                    //for (int k = 1; k <=gworkbook.Worksheets.Count; k++)
                    //{
                    //    Excel.Worksheet w = (Excel.Worksheet)gworkbook.Worksheets.get_Item(k);
                    //    if (w.Name.ToString() != "Comptes annuels" && w.Name.ToString() != "Annu.Refer")
                    //    {
                    //        w.Delete();
                    //    }
                    //}
                    myworkbook.Save();
                    enworkbook.Save();
                    gworkbook.Save();
                    deletesheets(myworkbook);
                    deletesheets(enworkbook);
                    deletesheets(gworkbook);

                    myworkbook.Close();
                    enworkbook.Close();
                    gworkbook.Close();
                    // releaseObject(deleteensheet);
                    releaseObject(enlanguesheet);
                    releaseObject(enworksheet);
                    releaseObject(enworkbook);
                    // releaseObject(deletegsheet);
                    releaseObject(glanguesheet);
                    releaseObject(gworksheet);
                    releaseObject(gworkbook);
                    // releaseObject(deletesheet);
                    releaseObject(deletelaguage);
                    releaseObject(myworksheet);
                    releaseObject(myworkbook);


                    app.Quit();

                }

            }

            //else if (checkBox23.Checked)
            //{

            //    string[] namestable = { "EVAL-SYNTHVALO2", "EVAL-SYNTHVALO1", "EVAL-SYNTHMULT1" };
            //    int rcont = 1;
            //    int ccont = 12;

            //    object misValue = System.Reflection.Missing.Value;
            //    for (int i = 0; i < namestable.Count(); i++)
            //    {
            //        string path = pathstylerfinal + "\\" + namestable[i] + ".xlsx";
            //        string enpath = "d:\\ptw\\notepme\\" + namestable[i] + "_EN.xlsx";
            //        string gpath = "d:\\ptw\\notepme\\" + namestable[i] + "_GER.xlsx";
            //        string frpath = "d:\\ptw\\notepme\\" + namestable[i] + "_FR.xlsx";

            //        Excel.Application app = new Excel.Application();
            //        app.DisplayAlerts = false;
            //        app.Visible = true;
            //        Excel.Workbook myworkbook;
            //        Excel.Workbook enworkbook;
            //        Excel.Workbook gworkbook;
            //        Excel.Worksheet myworksheet;
            //        Excel.Worksheet enworksheet;
            //        Excel.Worksheet gworksheet;
            //        Excel._Worksheet deleteworksheet1;
            //        Excel._Worksheet deleteworksheet2;

            //        myworkbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        // deleteworksheet1 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Histo.Macros-s");
            //        //  deleteworksheet2 = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Typologies IFRS-s");
            //        // deleteworksheet1.Delete();
            //        //deleteworksheet2.Delete();
            //        //Excel.Worksheet model = (Excel.Worksheet)myworkbook.Worksheets.get_Item("Modèles Goodwill");
            //        //model.Delete();
            //        myworkbook.SaveCopyAs(frpath);
            //        myworkbook.SaveCopyAs(enpath);
            //        myworkbook.SaveCopyAs(gpath);
            //        myworkbook.Close();
            //        myworkbook = app.Workbooks.Open(frpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        enworkbook = app.Workbooks.Open(enpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //        gworkbook = app.Workbooks.Open(gpath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //        myworksheet = (Excel.Worksheet)myworkbook.Worksheets.get_Item("SynthèseValorisations");
            //        enworksheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("SynthèseValorisations");
            //        gworksheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("SynthèseValorisations");


            //        //Excel.Worksheet enlanguesheet = (Excel.Worksheet)enworkbook.Worksheets.get_Item("Hist.Langues");
            //        //Excel.Worksheet glanguesheet = (Excel.Worksheet)gworkbook.Worksheets.get_Item("Hist.Langues");
            //        ////set using language to be english
            //        //Excel.Range enrange = enlanguesheet.get_Range("E4", "E1043");
            //        //Excel.Range enpasterange = enlanguesheet.get_Range("B4", "B1043");
            //        //enrange.Copy(enpasterange);
            //        //releaseObject(enrange);
            //        //releaseObject(enpasterange);
            //        ////set using language to be german
            //        //Excel.Range grange = glanguesheet.get_Range("F4", "f1043");
            //        //Excel.Range gpasterange = glanguesheet.get_Range("B4", "B1043");
            //        //grange.Copy(gpasterange);
            //        //releaseObject(grange);
            //        //releaseObject(gpasterange);

            //        //Excel.Range userange = myworksheet.UsedRange;
            //        //object[,] values = (object[,])userange.Value2;
            //        //Excel.Range copyrange;
            //        ////for (rcont = 1; rcont <= userange.Rows.Count; rcont++)
            //        ////{
            //        ////    string strcell = values[rcont, ccont].ToString();

            //        //try
            //        //{
            //        //    //fr 
            //        //    copyrange = myworksheet.get_Range(myworksheet.Cells[1, 1], myworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //        //    //en
            //        //    copyrange = enworksheet.get_Range(enworksheet.Cells[1, 1], enworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //        //    //german
            //        //    copyrange = gworksheet.get_Range(gworksheet.Cells[1, 1], gworksheet.Cells[userange.Rows.Count, 2]);
            //        //    copyrange.Copy(misValue);
            //        //    copyrange.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);


            //        //}
            //        //catch (Exception ex)
            //        //{
            //        //}
            //        //}
            //        //}

            //        //copyrange.Copy(copyrange);
            //        //Excel.Worksheet deletesheet =(Excel.Worksheet) myworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //Excel.Worksheet deleteensheet = (Excel.Worksheet)enworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //Excel.Worksheet deletegsheet = (Excel.Worksheet)gworkbook.Sheets.get_Item("Modèles Goodwill");
            //        //deletesheet.Delete();
            //        //deleteensheet.Delete();
            //        //deletegsheet.Delete();
            //        //Excel.Worksheet deletelaguage = (Excel.Worksheet)myworkbook.Sheets.get_Item("Hist.Langues");
            //        // deletelaguage.Delete();

            //        //enlanguesheet.Delete();
            //        // glanguesheet.Delete();


            //        myworkbook.Save();
            //        enworkbook.Save();
            //        gworkbook.Save();
            //        myworkbook.Close();
            //        enworkbook.Close();
            //        gworkbook.Close();

            //        //releaseObject(deleteworksheet1);
            //        // releaseObject(deleteworksheet2);
            //        // releaseObject(deleteensheet);
            //        // releaseObject(enlanguesheet);
            //        releaseObject(enworksheet);
            //        releaseObject(enworkbook);
            //        // releaseObject(deletegsheet);
            //        // releaseObject(glanguesheet);
            //        releaseObject(gworksheet);
            //        releaseObject(gworkbook);
            //        // releaseObject(deletesheet);
            //        //releaseObject(deletelaguage);
            //        releaseObject(myworksheet);
            //        releaseObject(myworkbook);

            //        app.Quit();

            //    }

        }
        private void hidden(string flag)
        {
            string path = "D:\\ptw\\notepme\\";
            string[] filename = Directory.GetFiles(path);
            Excel.Application xlapp = new Excel.ApplicationClass() as Excel.Application;
            xlapp.DisplayAlerts = false;
            xlapp.Application.DisplayAlerts = false;
            xlapp.Visible = true;

            for (int i = 0; i < filename.Length; i++)
            {
                string name = filename[i];
                if (!name.Contains("ANNUEL"))
                {
                    if (flag == "his")
                    {
                        Excel.Workbook xlworkbook = xlapp.Workbooks.Open(name);
                        Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Historique");
                        Excel.Range hide1 = xlworksheet.get_Range("N1", xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, xlworksheet.UsedRange.Columns.Count]);
                        hide1.EntireColumn.Hidden = true;

                        Excel.Range hide2 = xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1] as Excel.Range;
                        hide2.EntireRow.Hidden = true;

                        Excel.Range range1 = xlworksheet.get_Range("A1");
                        double height = double.Parse(range1.Height.ToString());

                        //for (int n = 1; n <= 4; n++)
                        //{
                        //    Excel.Range rangex = xlworksheet.get_Range("A" + n).EntireRow;
                        //    rangex.EntireRow.RowHeight = height /4;
                        //}


                        xlapp.ActiveWindow.DisplayGridlines = false;
                        xlworkbook.Save();
                        xlworkbook.Close();

                        releaseObject(xlworkbook);

                        releaseObject(xlworksheet);
                    }
                }
                else if (name.Contains("ANNUEL"))
                {
                    if (flag == "ann")
                    {
                        Excel.Workbook xlworkbook = xlapp.Workbooks.Open(name);
                        Excel.Worksheet xlworksheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item("Comptes annuels");
                        Excel.Range hide1 = xlworksheet.get_Range("Y1", xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, xlworksheet.UsedRange.Columns.Count]);
                        hide1.EntireColumn.Hidden = true;

                        Excel.Range hide2 = xlworksheet.Cells[xlworksheet.UsedRange.Rows.Count, 1] as Excel.Range;
                        hide2.EntireRow.Hidden = true;
                        xlapp.ActiveWindow.DisplayGridlines = false;
                        xlworkbook.Save();
                        xlworkbook.Close();

                        releaseObject(xlworkbook);

                        releaseObject(xlworksheet);
                    }
                }
            }
            xlapp.Quit();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            if (checkBox19.Checked)
            {
                historypagebreak(sender, e);
            }
            if (checkBox22.Checked)
            {
                companuelpagebreak(sender, e);
            }
        }
        private void sheetnamechange_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            //string findxo = "+Historique!C**Hist.Preface!D$14";
            object misValue = System.Reflection.Missing.Value;
            string paths = @"D:\ptw\prefaceNPS.xlsx";
            xlApp = new Excel.ApplicationClass();
            xlApp.DisplayAlerts = false;
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(paths, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            for (int nsx = 1; nsx <= xlWorkBook.Sheets.Count; nsx++)
            {
                Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(nsx);
                // admin1.UsedRange.Replace("Historique", "'Historique-s'", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                admin1.UsedRange.Replace("Hist.Preface", "'Hist.Preface-s'", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                admin1.UsedRange.Replace("Hist.Calculs", "'Hist.Calculs-s'", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                admin1.UsedRange.Replace("Hist.Langues", "'Hist.Langues-s'", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);
                admin1.UsedRange.Replace("Hist.Refer", "'Hist.Refer-s'", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, Type.Missing, false, false);

            }
            for (int nsx = 1; nsx <= xlWorkBook.Sheets.Count; nsx++)
            {
                Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(nsx);
                if (admin1.Name == "Historique")
                {
                    admin1.UsedRange.Clear();
                    admin1.UsedRange.ClearFormats();
                   // admin1.UsedRange.UseStandardWidth = true;
                }
                if (admin1.Name == "Hist.Preface")
                {
                    admin1.Name = "Hist.Preface-n";
                }
                if (admin1.Name == "Hist.Calculs")
                {
                    admin1.Name = "Hist.Calculs-n";
                }
                if (admin1.Name == "Hist.Langues")
                {
                    admin1.Name = "Hist.Langues-n";
                }
                if (admin1.Name == "Hist.Refer")
                {
                    admin1.Name = "Hist.Refer-n";
                }

            }

            for (int nsx = 1; nsx <= xlWorkBook.Sheets.Count; nsx++)
            {
                Excel.Worksheet admin1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(nsx);
                //if (admin1.Name == "Historique-s")
                //{
                //    admin1.Name = "Historique";
                //}
                if (admin1.Name == "Hist.Preface-s")
                {
                    admin1.Name = "Hist.Preface";
                }
                if (admin1.Name == "Hist.Calculs-s")
                {
                    admin1.Name = "Hist.Calculs";
                }
                if (admin1.Name == "Hist.Langues-s")
                {
                    admin1.Name = "Hist.Langues";
                }
                if (admin1.Name == "Hist.Refer-s")
                {
                    admin1.Name = "Hist.Refer";
                }

            }
            Excel.Worksheet sheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
          
            Excel.Worksheet sheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
           // MessageBox.Show("rows:"+sheet2.Rows.Count+" columns:"+sheet2.Columns.Count);
         
            sheet2.UsedRange.Cut(sheet1.UsedRange);
            sheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            int countr = sheet1.UsedRange.Rows.Count;
            int countc = sheet1.UsedRange.Columns.Count;
            object[,] values = (object[,])sheet1.UsedRange.Value2;
            int contendc = 0;
            int contendr = 0;
            for (int i = 1; i < countc; i++)
            {
                if (values[1, i] != null)
                {
                    if (values[1, i].ToString() != "" && values[2, i] != null && values[2, i].ToString() != "" && values[1, i].ToString().Trim() == "1000" && values[2, i].ToString().Trim() == "2000")
                    {
                        contendc = i;
                        break;
                    }
                }
            }
            for (int i = 1; i < countr; i++)
            {
                if (values[i, 1] != null)
                {
                    if (values[i, 1].ToString() != "" && values[i, 2] != null && values[i, 2].ToString() != "" && values[i, 1].ToString().Trim() == "1000" && values[i, 2].ToString().Trim() == "2000")
                    {
                        contendr = i;
                        break;
                    }
                }
            }

         //   sheet1.get_Range(sheet1.Cells[1, contendc+1], sheet1.Cells[1, countc]).EntireColumn.Cells.Delete();


         //   sheet1.get_Range(sheet1.Cells[contendr+1, 1], sheet1.Cells[countr, 1]).EntireRow.Cells.Delete();
                sheet2.Delete();

            xlWorkBook.SaveAs(@"D:\ptw\prefaceNPS.xlsx");
            xlWorkBook.Close();
            xlApp.Quit();
        }
        private void simply_buttonlancer_Click(object sender, EventArgs e)
        {
            simply_leger_Click(sender, e);
            sheetnamechange_Click(sender, e);
            MessageBox.Show("Opérations terminées " + timleger + " secondes");
        }
        private void simply_leger_Click(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;

            fichierprepare = textBox11.Text;
            prefaceNP = "D:\\ptw\\prefaceNP.xlsx";

            fichierprepare = "D:\\ptw\\prefaceNPS.xlsx";
            prefaceNP = "D:\\ptw\\prefaceNPS.xlsx";

            supprimerTypologie_Click(sender, e);

            button2_Click(sender, e);
            HistoCalculs();

            HistoMettreZero_Click(sender, e);
            HistoRempl_Click(sender, e);

            HistoAuAvAw_Click(sender, e);
            colCE_Click(sender, e);//72000
            supprimerREF_Click(sender, e);

            ////////////Histo.ptw et histo.preface
            button1_Click(sender, e);//Inserer les colonnes correctifs
            Histopreface_Click(sender, e);

            ////////Annuel .ptw
            AnnuelO_Click(sender, e);
            //ComptesAnnuels_Click(sender, e);

            supprimercol_Click(sender, e);
            button5_Click(sender, e);//supprimer ligne -1

            //supprimer les onglets
            Supprimeronglet_Click(sender, e);

            //traitement REF!
            Historique84000();
            //fonctionRemplacerD1();//D1 formule trop longue
            consigneProteger();

            //Pour Hist-s legement
            insertionHistoS(sender, e);//Inserer les colonnes correctifs
            insertColumnHistoS();
            HistoprefaceHistoS(sender, e);//inserer les colonnes pour Hist.Preface-s
            supprimercolhistoS(sender, e);
            consigneProtegerHistoS();


            changerNumeroligne();//1000-100 pour historique //Param Sav mettre a "1" ---- OLEDB.net
            copyrangeannewlrefer(sender, e);

          //  button1_Click(sender, e);//Inserer les colonnes correctifs
           // Histopreface_Click(sender, e);




           

            int time2 = System.Environment.TickCount;
            int times = (time2 - time1) / 1000;

            int hours = times / 3600;
            int minuit = times / 60 - hours * 60;
            int second = times - minuit * 60 - hours * 3600;
            timleger = hours + " heures " + minuit + " minutes " + second;
            //timleger = Convert.ToString(Convert.ToDecimal(times) / 1000);
        }
        private void insertColumnHistoS()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
           // prefaceNP = "D:\\ptw\\prefaceNP.xlsx";
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            Excel.Worksheet xlworksheet_s = xlWorkBook.Worksheets["Historique-s"] as Excel.Worksheet;
            CodeFinder cf;
            cf = new CodeFinder(xlWorkBook, xlworksheet_s);
            Excel.Range range_s = xlworksheet_s.UsedRange;
            Excel.Range rangex4_s = xlworksheet_s.Cells[1, 3] as Excel.Range;
            string insert41_s = cf.FindCodedColumnHeader("3000", range_s);

            rangex4_s.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            Excel.Range rangex5_s = xlworksheet_s.Cells[1, 3] as Excel.Range;
            string insert51_s = cf.FindCodedColumnHeader("3000", range_s);
            rangex5_s.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            //string insert4_s = cf.FindCodedColumnHeader("82000-1000", range_s);
            //string insert5_s = cf.FindCodedColumnHeader("82000-2000", range_s);
            //Excel.Range rangex2c_s = xlworksheet_s.UsedRange.get_Range(insert4_s + "1", insert5_s + "1") as Excel.Range;
            //rangex2c_s.EntireColumn.Copy(xlworksheet_s.UsedRange.get_Range(insert41_s + "1", insert51_s + "1").EntireColumn);


            xlWorkBook.Save();

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlworksheet_s);
            releaseObject(xlWorkBook);


        }

        private void LockStateFiles(object sender, EventArgs e)
        {
            try
            {
                textBox20.AppendText("==> Start Protection des cellules" + System.Environment.NewLine);
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                int time1 = System.Environment.TickCount;
                object misValue = System.Reflection.Missing.Value;
                prefaceNP = "D:\\ptw\\prefaceNPS.xlsx";
                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


                if (checkBox24.Checked)
                {
                    Excel.Worksheet xlWorkSheetX = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                    LockStateFiles(textBox17.Text + "\\lockedStatus3.stat", xlWorkSheetX);
                }
                if (checkBox25.Checked)
                {
                    Excel.Worksheet xlWorkSheetX = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Comptes annuels");
                    LockStateFile(textBox17.Text + "\\lockedStatus4.stat", xlWorkSheetX);
                }
                int time12 = System.Environment.TickCount;
                int time = ((time12 - time1) / 1000);
                xlWorkBook.Close();
                xlApp.Quit();
                int hours = time / 3600;
                int minuit = time / 60 - hours * 60;
                int second = time - minuit * 60 - hours * 3600;
                string timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Protection des cellules OK : " + timeto + " s" + System.Environment.NewLine);
                //  MessageBox.Show("Protection des cellules OK : " + timeto + " s");


                textBox20.AppendText("==> Start Raz PrefaceNP. Le fichier sera sauvé dans D:\\ptw\\notepme !" + System.Environment.NewLine);
                int timex = System.Environment.TickCount;
                FileStream file = new FileStream(textBox17.Text + "\\lockedStatus3.stat", FileMode.Open, FileAccess.Read);
                List<String> lockchecklist;
                lockchecklist = getLockedstatusList(file);
                // Excel.Application xlApp;
                //   Excel.Workbook xlWorkBook;

                // object misValue = System.Reflection.Missing.Value;
                prefaceNP = "D:\\ptw\\prefaceNPS.xlsx";
                xlApp = new Excel.ApplicationClass();
                xlApp.Visible = true;
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
                //Excel.Range usedrange = xlWorkSheet.UsedRange;
                object[,] values = (object[,])xlWorkSheet.UsedRange.Value2;
                object[,] formulas = (object[,])xlWorkSheet.UsedRange.Formula;
               // int rowcount = 654;
               // int columncount = 63;
                int rowcount = xlWorkSheet.get_Range("A1", "A661").Rows.Count-1;
                int columncount = xlWorkSheet.get_Range("A1", "BK1").Columns.Count;
                int rowp = 0;
                int rowp2 = 0;
                 for (int i = 1; i < columncount - 1; i++)
                    {
                        if (values[rowcount, i] != null)
                        {
                            if (values[rowcount, i].ToString() == "83000")
                            {
                                rowp = i;
                            }
                            if (values[rowcount, i].ToString() == "84000")
                            {
                                rowp2 = i;
                                break;
                            }


                        }
                    }
                 for (int i = 1; i < rowcount - 1; i++)
                 {

                     for (int j = 1; j < columncount - 2; j++)
                     {
                         if (values[i, j] != null)
                         {
                             if (j == rowp || j == rowp2)
                             {
                                 string xss = "ss";
                             }
                             if (i == 19 || i == 20 || i == 21)
                             {
                             }
                             else
                             {
                                 if (getLockedstatus(lockchecklist, i, j) == false)
                                 {

                                     if (values[i, columncount] != null)
                                     {
                                         //string v = values[i, columncount].ToString();
                                         //if (values[i, columncount].ToString() == "224000" || values[i, columncount].ToString() == "242000-12000" || values[i, columncount].ToString() == "762000-4000" || values[i, columncount].ToString() == "763000-2000" || values[i, columncount].ToString() == "763000-4000" || values[i, columncount].ToString() == "243000" || values[i, columncount].ToString() == "243000-1000" || values[i, columncount].ToString() == "299000-200" || values[i, columncount].ToString() == "468000" || values[i, columncount].ToString() == "471000" || values[i, columncount].ToString() == "473000-1000" || values[i, columncount].ToString() == "475000" || values[i, columncount].ToString() == "478000" || values[i, columncount].ToString() == "480000-1000" || values[i, columncount].ToString() == "482000" || values[i, columncount].ToString() == "485000" || values[i, columncount].ToString() == "487000-1000" || values[i, columncount].ToString() == "745000-2000" || values[i, columncount].ToString() == "745000-3000" || values[i, columncount].ToString() == "746000-2000" || values[i, columncount].ToString() == "746000-3000" || values[i, columncount].ToString() == "768000" || values[i, columncount].ToString() == "772000" || values[i, columncount].ToString() == "776000" || values[i, columncount].ToString() == "780000" || values[i, columncount].ToString() == "791000-500" || values[i, columncount].ToString() == "791000-700" || values[i, columncount].ToString() == "791000-1000" || values[i, columncount].ToString() == "814000-1000" || values[i, columncount].ToString() == "816000-1000" || values[i, columncount].ToString() == "818000-1000" || values[i, columncount].ToString() == "308000" || j == rowp || j == rowp2)
                                         //{
                                         //    values[i, j] = formulas[i, j];
                                         //}
                                         //else
                                         //{
                                             values[i, j] = 0;
                                        // }
                                     }
                                     else
                                     {
                                         //values[i, j] = formulas[i, j];
                                     }
                                 }
                                 else
                                 {
                                     values[i, j] = formulas[i, j];
                                 }
                             }
                         }
                     }

                 }

                
                xlWorkSheet.UsedRange.Formula = values;

                xlWorkBook.SaveCopyAs("D:\\ptw\\notepme\\prefaceNPS.xlsx");
                xlWorkBook.Close();
                xlApp.Quit();
                int timey = System.Environment.TickCount;
                int x = (timey - timex) / 1000;
                hours = x / 3600;
                minuit = x / 60 - hours * 60;
                second = x - minuit * 60 - hours * 3600;
                timeto = hours.ToString() + " heures " + minuit.ToString() + " minutes " + second.ToString();
                textBox20.AppendText("Raz PrefaceNP OK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s" + System.Environment.NewLine);
                if (flagsimplypastout)
                {
                    MessageBox.Show("Raz PrefaceNPOK. Le fichier est sauvé dans D:\\ptw\\notepme ! Time : " + timeto + "s");
                }
                else
                {
                    flagsimplypastout = true;
                }


            }
            catch (Exception ex)
            {
                textBox20.AppendText(ex.ToString() + System.Environment.NewLine);
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            Process[] myprocess = Process.GetProcesses();


            bool pcheck1 = false;

            
                Process[] MyProcess = Process.GetProcessesByName("autoupdatepack");
                foreach (Process p1 in MyProcess)
                {
                    string sss = p1.StartTime.ToString();
                }
                foreach (Process p1 in myprocess)
                {
                    try
                    {
                        string processName = p1.ProcessName.ToLower().Trim();
                        if (processName == "autoupdatepack")
                        {
                            pcheck1 = true;

                        }

                    }
                    catch { }
                }

                if (!pcheck1)
                {
                    Process pro = new Process();
                    string FileName = @"D:\Alex\Transformer Fichier EXCEL\autoupdatepack.exe";
                    if (System.IO.File.Exists(FileName))
                    {
                        runprogram(FileName, "", "");


                        pcheck1 = true;

                    }
                }

            
        }
        public void runprogram(string programname, string cmd, string wdirectory)
        {
            Process proc = new Process();
            //proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.FileName = programname;
            proc.StartInfo.Arguments = cmd;
            proc.StartInfo.UseShellExecute = false;
            if (wdirectory != "")
                proc.StartInfo.WorkingDirectory = wdirectory;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.RedirectStandardInput = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.Start();
        }
    }
}