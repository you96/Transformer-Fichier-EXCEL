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


namespace TransformEXCEL
{
    class HistoSimplifier
    {
        string fichierprepare = null;
        string prefaceNP = null;

        //
        //// Supprimer typologie dans "D:\\ptw\\Histo.ptw"
        //
        private void renommer()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            //missing values, for example, when you invoke methods that have default parameter values.
            //remplace les paramètres par default des fonctions utilisés

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(fichierprepare, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            //Afficher pas les Alerts !!non utiliser avant assurer!!!
            xlApp.DisplayAlerts = false;

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique");
            //Excel.Worksheet sheetTypologie = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Typologie IFRS");
            //sheetTypologie.Delete();

            xlWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);//kill WorkSheet
            releaseObject(xlWorkBook);//Kill WorkBook
            releaseObject(xlApp);//Kill Application Application 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            CodeFinder cf = new CodeFinder(prefaceNP, "Historique-s");
           /* xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");*/

            Excel.Range range = cf.XlsWorkSheet.UsedRange;

            Excel.Range rangex1 = cf.XlsWorkSheet.Cells[1, 4] as Excel.Range;

            Excel.Range rangex2 = cf.XlsWorkSheet.Cells[1, 5] as Excel.Range;

            Excel.Range rangex3 = cf.XlsWorkSheet.Cells[1, 6] as Excel.Range;


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
            Excel.Range rangex1c = cf.XlsWorkSheet.UsedRange.get_Range("EU1", "EV1") as Excel.Range;
            rangex1c.EntireColumn.Copy(cf.XlsWorkSheet.UsedRange.get_Range("D1", "E1").EntireColumn);
            rangex1c.EntireColumn.Copy(cf.XlsWorkSheet.UsedRange.get_Range("G1", "H1").EntireColumn);
            rangex1c.EntireColumn.Copy(cf.XlsWorkSheet.UsedRange.get_Range("J1", "K1").EntireColumn);


            Excel.Worksheet xlWorkSheet2 = cf.XlsWorkBook.Worksheets["Hist.Refer"] as Excel.Worksheet;
            Excel.Range rangeC = xlWorkSheet2.Cells[1, 4] as Excel.Range;

            Excel.Range rangeD = xlWorkSheet2.Cells[1, 5] as Excel.Range;

            Excel.Range rangeE = xlWorkSheet2.Cells[1, 6] as Excel.Range;

            rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangeC.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangeD.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

            rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);
            rangeE.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, misValue);

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
            rangeDc.EntireColumn.Copy(rangeDc2.EntireColumn);
            rangeDc.EntireColumn.Copy(rangeDc3.EntireColumn);
            rangeEc.EntireColumn.Copy(rangeEc2.EntireColumn);
            rangeEc.EntireColumn.Copy(rangeEc3.EntireColumn);

            //Excel.Worksheet WorkSheetPreface = xlWorkBook.Worksheets["Hist.Preface"] as Excel.Worksheet;


            Excel.Range rangex1cx1 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 4] as Excel.Range;
            Excel.Range rangex1cx2 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 5] as Excel.Range;
            Excel.Range rangex2cx1 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 7] as Excel.Range;
            Excel.Range rangex2cx2 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 8] as Excel.Range;
            Excel.Range rangex3cx1 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 10] as Excel.Range;
            Excel.Range rangex3cx2 = cf.XlsWorkSheet.Cells[range.Rows.Count - 1, 11] as Excel.Range;

            rangex1cx1.Value2 = "";
            rangex1cx2.Value2 = "";
            rangex2cx1.Value2 = "";
            rangex2cx2.Value2 = "";
            rangex3cx1.Value2 = "";
            rangex3cx2.Value2 = "";




            //tester EE pour Histo.refer//et parcourir historique

            Excel.Range rangeRefer = xlWorkSheet2.UsedRange;
            Excel.Range rangeHistorique = cf.XlsWorkSheet.UsedRange;
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
                            Excel.Range cellcopie = cf.XlsWorkSheet.Cells[rowHistoCnt, 3] as Excel.Range;
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 4]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 5]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 6]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 7]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 8]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 9]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 10]);
                            cellcopie.Copy(cf.XlsWorkSheet.Cells[rowHistoCnt, 11]);
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
            cf.XlsWorkSheet.SaveAs(prefaceNP, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            cf.XlsApp.DisplayAlerts = true;
            cf.XlsWorkBook.Close(true, misValue, misValue);
            cf.XlsApp.Quit();

            //MessageBox.Show("jobs done!");
            releaseObject(cf.XlsWorkSheet);
            releaseObject(cf.XlsWorkBook);
            releaseObject(cf.XlsApp);
        }



        private void supprimercol_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int time1 = System.Environment.TickCount;
            ////////////////////////////////////////400000//////////////////////
            int rCnt = 0;
            int cCnt = 0;
            int row400000 = 0;

            cCnt = range.Columns.Count;
            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "400000"))
                {
                    row400000 = rCnt;
                    break;
                }
            }

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

        private void consigneProteger()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = true;
            xlApp.DisplayAlerts = false;

            xlWorkBook = xlApp.Workbooks.Open(prefaceNP, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, true, false);

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item("Historique-s");
            Excel.Range range = xlWorkSheet.UsedRange;
            int rowcount = xlWorkSheet.UsedRange.Rows.Count;
            object[,] values = (object[,])range.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            rCnt = range.Rows.Count;
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "15000"))
                {
                    col = cCnt;
                    break;
                }
            }

            //Routine pour modifier col XXXXX marquer ligne proteger -1
            for (int i = 1; i < rowcount - 5; i++)
            {
                if ((xlWorkSheet.Cells[i, 3] as Excel.Range).Locked.ToString() == "True")
                    (xlWorkSheet.Cells[i, col] as Excel.Range).Value2 = "-1";
            }


            xlApp.Save(misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        public void leger_Click(object sender, EventArgs e)
        {
            int time1 = System.Environment.TickCount;

            fichierprepare = "D:\\ptw\\preface.xls";
            prefaceNP = "D:\\ptw\\prefaceNP.xls";

            renommer();

            //button2_Click(sender, e);
            //HistoCalculs();
            //HistoMettreZero_Click(sender, e);
            //HistoRempl_Click(sender, e);
            //HistoAuAvAw_Click(sender, e);
            //colCE_Click(sender, e);//72000
            //supprimerREF_Click(sender, e);

            ////////////Histo.ptw et histo.preface
            button1_Click(sender, e);//Inserer les colonnes correctifs
            //Histopreface_Click(sender, e);

            ////////Annuel .ptw
            //AnnuelO_Click(sender, e);
            //ComptesAnnuels_Click(sender, e);
            supprimercol_Click(sender, e);

            //button5_Click(sender, e);

            //supprimer les onglets
            //Supprimeronglet_Click(sender, e);

            //traitement REF!
            //Historique84000();
            //fonctionRemplacerD1();
            consigneProteger();



            int time2 = System.Environment.TickCount;
            int times = time2 - time1;
            string timleger = Convert.ToString(Convert.ToDecimal(times) / 1000);
            MessageBox.Show(timleger);

            }

        public static void Callleger_Click(object sender, EventArgs e)
        {

        }

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
    }
}
