using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace TransformEXCEL
{
    public partial class FicherUpdate : Form
    {
        public FicherUpdate()
        {
            InitializeComponent();
        }
        public void backupCopyfiles(string[] filepath, string path, string[] filepath2,string type)
        {
            if (type == "1")
            {
                if (System.Environment.MachineName.ToUpper() == "PATRICK-6800")
                {
                    if (path.ToUpper().Contains("PATRICK-6800"))
                    {
                        path = path.Replace("\\\\PATRICK-6800\\d\\", "d:\\");
                        path = path.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                        path = path.Replace(@"\\\\PATRICK-6800\\d\\", "d:\\");
                        path = path.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                    }
                    if (!Directory.Exists(path + "backup"))
                    {
                        Directory.CreateDirectory(path + "backup");
                    }
                }
                else
                {
                    if (!Directory.Exists(path + "backup"))
                    {
                        Directory.CreateDirectory(path + "backup");
                    }
                }
                
                for (int x = 0; x < filepath.Length; x++)
                {
                    // Copy the file and increment the ProgressBar if successful.
                    CopyFile(filepath[x], filepath2[x]);
                }
            }
            else
            {
                if (Directory.Exists(path + "backup"))
                {

                    for (int x = 0; x < filepath.Length; x++)
                    {
                        // Copy the file and increment the ProgressBar if successful.
                        CopyFile(filepath2[x], filepath[x]);
                    }
                }
            }

        }
        public void backup(string type)
        {

            string[] filenamebothindex = { "Historique_Rows.index", "Historique_Columns.index", "Historique-s_Rows.index", "Comptes annuels_Rows.index", "Historique-s_Columns.index", "Comptes annuels_Columns.index", "CmpcWacc_Rows.index", "CalculFCF_Rows.index", "DiscountedFCF_Rows.index", "MéthodesMixtes_Rows.index", "Multiples_Rows.index", "TransactionsComparables_Rows.index", "AutresCapitalisations_Rows.index", "GordonShapiroBates_Rows.index", "Goodwill_Rows.index", "PatrimonialAncAncc_Rows.index", "CmpcWacc_Columns.index", "CalculFCF_Columns.index", "DiscountedFCF_Columns.index", "MéthodesMixtes_Columns.index", "Multiples_Columns.index", "TransactionsComparables_Columns.index", "AutresCapitalisations_Columns.index", "GordonShapiroBates_Columns.index", "Goodwill_Columns.index", "PatrimonialAncAncc_Columns.index" };
            string[] filenamebothNP = { "prefaceNP.xlsx", "prefaceNPS.xlsx" };
            string[] filenamebothcell = { "lockedStatus.stat", "lockedStatus2.stat", "lockedStatus3.stat" };
            string[] filenamebothsamll = { "S-ACT", "S-PAS", "S-CR", "S-ANN5", "S-ANN4", "S-ANN3", "ACT1", "ACT4", "PAS1", "PAS3", "CR1", "CR3", "ANN5-1", "ANN5-2", "ANN6-1", "ANN6-2", "ANN6-3", "ANN7-1", "ANN7-2", "ANN7-3", "ANN8-1", "ANN8-2", "ANN11-1", "ANNUEL-CR1", "ANNUEL-CR2", "ANNUEL-CR3", "ANNUEL-BILACT1", "ANNUEL-BILACT1", "ANNUEL-BILPAS1", "ANNUEL-BILUSACT1", "ANNUEL-BILUSPAS1", "ANNUEL-FLUXFIN1", "ANNUEL-FLUXTRES1", "ANNUEL-RATIOS1", "ANNUEL-RATIOS2", "ANNUEL-SYNTH", "BilanIFRS", "CRIFRS", "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc" };

         
            string path1 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
            
            if (radioButton2.Checked)
            {
                path1 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
            }
            else if (radioButton1.Checked)
            {
                path1 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
            }
            else if (radioButton3.Checked)
            {
                path1 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
            }
            else if (radioButton4.Checked)
            {
                path1 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
            }
            string text = "";
            for (int i = 0; i < filenamebothNP.Length; i++)
            {
                text = text + path1 + filenamebothNP[i] + "|";

            }
            string[] filenames = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames = new string[] { text.Substring(0, text.Length - 1) };
            }
            text = "";
            for (int i = 0; i < filenamebothNP.Length; i++)
            {
                text = text + path1+"backup\\" + filenamebothNP[i] + "|";

            }
            string[] filenames2 = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames2 = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames2 = new string[] { text.Substring(0, text.Length - 1) };
            }
            backupCopyfiles(filenames, path1, filenames2,type);

            //index
            if (radioButton2.Checked)
            {
                path1 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
            }
            else if (radioButton1.Checked)
            {
                path1 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
            }
            else if (radioButton3.Checked)
            {
                path1 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
            }
            else if (radioButton4.Checked)
            {
                path1 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
            }
            text = "";
            for (int i = 0; i < filenamebothindex.Length; i++)
            {
                text = text + path1 + filenamebothindex[i] + "|";

            }
            filenames = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames = new string[] { text.Substring(0, text.Length - 1) };
            }
            text = "";
            for (int i = 0; i < filenamebothindex.Length; i++)
            {
                text = text + path1 + "backup\\" + filenamebothindex[i] + "|";
            }
            filenames2 = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames2 = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames2 = new string[] { text.Substring(0, text.Length - 1) };
            }
            backupCopyfiles(filenames, path1, filenames2,type);
            //samllfile

            if (radioButton2.Checked)
            {
                path1 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
            }
            else if (radioButton1.Checked)
            {
                path1 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
            }
            else if (radioButton3.Checked)
            {
                path1 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
            }
            else if (radioButton4.Checked)
            {
                path1 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
            }
            text = "";

            string[] tempStrings = { "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc" };
            for (int i = 0; i < filenamebothsamll.Length; i++)
            {
              
                if (tempStrings.Contains(filenamebothsamll[i]))
                {
                    text = text + path1 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                }
                else
                {
                    text = text + path1 + filenamebothsamll[i] + "_EN.xlsx" + "|";
                    text = text + path1 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                    text = text + path1 + filenamebothsamll[i] + "_GER.xlsx" + "|";
                }
            }
            
          
            filenames = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames = new string[] { text.Substring(0, text.Length - 1) };
            }
            text = "";
            for (int i = 0; i < filenamebothsamll.Length; i++)
            {
                
                if (tempStrings.Contains(filenamebothsamll[i]))
                {
                    text = text + path1  + "backup\\"+ filenamebothsamll[i] + "_FR.xlsx" + "|";
                }
                else
                {
                    text = text + path1 + "backup\\" + filenamebothsamll[i] + "_EN.xlsx" + "|";
                    text = text + path1 + "backup\\" + filenamebothsamll[i] + "_FR.xlsx" + "|";
                    text = text + path1 + "backup\\" + filenamebothsamll[i] + "_GER.xlsx" + "|";
                }
            }
            filenames2 = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames2 = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames2 = new string[] { text.Substring(0, text.Length - 1) };
            }
            backupCopyfiles(filenames, path1, filenames2,type);

            //pilotage
            if (radioButton2.Checked)
            {
                path1 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
            }
            else if (radioButton1.Checked)
            {
                path1 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
            }
            else if (radioButton3.Checked)
            {
                path1 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
            }
            else if (radioButton4.Checked)
            {
                path1 = @"\\Francis-7500\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
            }
            text = "";
            for (int i = 0; i < filenamebothindex.Length; i++)
            {
                text = text + path1 + filenamebothindex[i] + "|";

            }
            filenames = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames = new string[] { text.Substring(0, text.Length - 1) };
            }
            text = "";
            for (int i = 0; i < filenamebothindex.Length; i++)
            {
                text = text + path1 + "backup\\" + filenamebothindex[i] + "|";
            }
            filenames2 = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames2 = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames2 = new string[] { text.Substring(0, text.Length - 1) };
            }
            backupCopyfiles(filenames, path1, filenames2,type);
            text = "";
            for (int i = 0; i < filenamebothcell.Length; i++)
            {
                text = text + path1 + filenamebothcell[i] + "|";

            }
            filenames = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames = new string[] { text.Substring(0, text.Length - 1) };
            }
            text = "";
            for (int i = 0; i < filenamebothcell.Length; i++)
            {
                text = text + path1 + "backup\\" + filenamebothcell[i] + "|";
            }
            filenames2 = text.Split('|');
            if (text.Substring(0, text.Length - 1).Contains("|"))
            {
                filenames2 = text.Substring(0, text.Length - 1).Split('|');
            }
            else
            {
                filenames2 = new string[] { text.Substring(0, text.Length - 1) };
            }
            backupCopyfiles(filenames, path1, filenames2,type);
        }
        private void button1_Click(object sender, EventArgs e)
        {
           
            bool flagLoad = true;
            DialogResult dr = MessageBox.Show("Do you want to back up all the files before update the new files", "Information", MessageBoxButtons.YesNoCancel);
            if ( dr== DialogResult.Yes)
            {
                try
                {
                    backup("1");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }else if ( dr== DialogResult.No){
            }else{
                flagLoad = false;
            }
            if (flagLoad)
            {
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                string[] filenamessimplyindex = { "Historique-s_Rows.index", "Comptes annuels_Rows.index", "Historique-s_Columns.index", "Comptes annuels_Columns.index", "P_Rows.index", "P_Columns.index", "CmpcWacc_Rows.index", "CalculFCF_Rows.index", "DiscountedFCF_Rows.index", "MéthodesMixtes_Rows.index", "Multiples_Rows.index", "TransactionsComparables_Rows.index", "AutresCapitalisations_Rows.index", "GordonShapiroBates_Rows.index", "Goodwill_Rows.index", "PatrimonialAncAncc_Rows.index", "CmpcWacc_Columns.index", "CalculFCF_Columns.index", "DiscountedFCF_Columns.index", "MéthodesMixtes_Columns.index", "Multiples_Columns.index", "TransactionsComparables_Columns.index", "AutresCapitalisations_Columns.index", "GordonShapiroBates_Columns.index", "Goodwill_Columns.index", "PatrimonialAncAncc_Columns.index", "ChoixMéthodes_Columns.index", "ChoixMéthodes_Rows.index", "SynthèseValorisations_Columns.index", "SynthèseValorisations_Rows.index", "Hist.Refer_Rows.index", "Hist.Refer_Columns.index" };
                string[] filenamesnormarlindex = { "Historique_Rows.index", "Comptes annuels_Rows.index", "Historique_Columns.index", "Comptes annuels_Columns.index", "P_Rows.index", "P_Columns.index", "CmpcWacc_Rows.index", "CalculFCF_Rows.index", "DiscountedFCF_Rows.index", "MéthodesMixtes_Rows.index", "Multiples_Rows.index", "TransactionsComparables_Rows.index", "AutresCapitalisations_Rows.index", "GordonShapiroBates_Rows.index", "Goodwill_Rows.index", "PatrimonialAncAncc_Rows.index", "CmpcWacc_Columns.index", "CalculFCF_Columns.index", "DiscountedFCF_Columns.index", "MéthodesMixtes_Columns.index", "Multiples_Columns.index", "TransactionsComparables_Columns.index", "AutresCapitalisations_Columns.index", "GordonShapiroBates_Columns.index", "Goodwill_Columns.index", "PatrimonialAncAncc_Columns.index", "ChoixMéthodes_Columns.index", "ChoixMéthodes_Rows.index", "SynthèseValorisations_Columns.index", "SynthèseValorisations_Rows.index", "Hist.Refer_Rows.index", "Hist.Refer_Columns.index", "block.index" };
                string[] filenamebothindex = { "Historique_Rows.index", "Historique_Columns.index", "Historique-s_Rows.index", "Comptes annuels_Rows.index", "P_Rows.index", "P_Columns.index", "CmpcWacc_Rows.index", "CalculFCF_Rows.index", "DiscountedFCF_Rows.index", "MéthodesMixtes_Rows.index", "Multiples_Rows.index", "TransactionsComparables_Rows.index", "AutresCapitalisations_Rows.index", "GordonShapiroBates_Rows.index", "Goodwill_Rows.index", "PatrimonialAncAncc_Rows.index", "CmpcWacc_Columns.index", "CalculFCF_Columns.index", "DiscountedFCF_Columns.index", "MéthodesMixtes_Columns.index", "Multiples_Columns.index", "TransactionsComparables_Columns.index", "AutresCapitalisations_Columns.index", "GordonShapiroBates_Columns.index", "Goodwill_Columns.index", "PatrimonialAncAncc_Columns.index", "ChoixMéthodes_Columns.index", "ChoixMéthodes_Rows.index", "SynthèseValorisations_Columns.index", "SynthèseValorisations_Rows.index", "Hist.Refer_Rows.index", "Hist.Refer_Columns.index" };
                string[] filenamessimplyNP = { "prefaceNPS.xlsx" };
                string[] filenamesnormarNP = { "prefaceNP.xlsx" };
                string[] filenamebothNP = { "prefaceNP.xlsx", "prefaceNPS.xlsx" };
                string[] filenamessimplycell = { "lockedStatus3.stat" };
                string[] filenamesnormarcell = { "lockedStatus.stat", "lockedStatus2.stat" };
                string[] filenamebothcell = { "lockedStatus.stat", "lockedStatus2.stat", "lockedStatus3.stat" };
                string[] filenamessimplysamll = { "S-ACT", "S-PAS", "S-CR", "S-ANN5", "S-ANN4", "S-ANN3", "ANNUEL-CR1", "ANNUEL-CR2", "ANNUEL-CR3", "ANNUEL-BILACT1", "ANNUEL-BILACT1", "ANNUEL-BILPAS1", "ANNUEL-BILUSACT1", "ANNUEL-BILUSPAS1", "ANNUEL-FLUXFIN1", "ANNUEL-FLUXTRES1", "ANNUEL-RATIOS1", "ANNUEL-RATIOS2", "ANNUEL-SYNTH" };
                string[] filenamesnormarsamll = { "ACT1", "ACT4", "PAS1", "PAS3", "CR1", "CR3", "ANN5-1", "ANN5-2", "ANN6-1", "ANN6-2", "ANN6-3", "ANN7-1", "ANN7-2", "ANN7-3", "ANN8-1", "ANN8-2", "ANN11-1", "ANNUEL-CR1", "ANNUEL-CR2", "ANNUEL-CR3", "ANNUEL-BILACT1", "ANNUEL-BILACT1", "ANNUEL-BILPAS1", "ANNUEL-BILUSACT1", "ANNUEL-BILUSPAS1", "ANNUEL-FLUXFIN1", "ANNUEL-FLUXTRES1", "ANNUEL-RATIOS1", "ANNUEL-RATIOS2", "ANNUEL-SYNTH", "BilanIFRS", "CRIFRS", "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc", "APNNE" };
                string[] filenamebothsamll = { "S-ACT", "S-PAS", "S-CR", "S-ANN5", "S-ANN4", "S-ANN3", "ACT1", "ACT4", "PAS1", "PAS3", "CR1", "CR3", "ANN5-1", "ANN5-2", "ANN6-1", "ANN6-2", "ANN6-3", "ANN7-1", "ANN7-2", "ANN7-3", "ANN8-1", "ANN8-2", "ANN11-1", "ANNUEL-CR1", "ANNUEL-CR2", "ANNUEL-CR3", "ANNUEL-BILACT1", "ANNUEL-BILACT1", "ANNUEL-BILPAS1", "ANNUEL-BILUSACT1", "ANNUEL-BILUSPAS1", "ANNUEL-FLUXFIN1", "ANNUEL-FLUXTRES1", "ANNUEL-RATIOS1", "ANNUEL-RATIOS2", "ANNUEL-SYNTH", "APNNE" };

                int flag = 0;
                if (comboBox1.Text == "Simply")
                {

                    flag = 1;
                }
                else if (comboBox1.Text == "Normal")
                {
                    flag = 2;
                }
                else if (comboBox1.Text == "Both")
                {
                    flag = 3;
                }
                else
                {
                    MessageBox.Show("Please choose one source: simply, normal or both!");
                }
                if (flag != 0)
                {
                    if (radioButton1.Checked || radioButton2.Checked || radioButton3.Checked || radioButton4.Checked)
                    {
                        string path1 = @"d:\ptw\notepme\";
                        string path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                        string path3 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                        if (flag == 1)
                        {
                            //prefaceNP
                            string text = "";

                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            for (int i = 0; i < filenamessimplyNP.Length; i++)
                            {
                                text = text + path1 + filenamessimplyNP[i] + "|";

                            }
                            string[] filenames = text.Split('|');

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }

                            text = "";
                            for (int i = 0; i < filenamessimplyNP.Length; i++)
                            {
                                text = text + path2 + filenamessimplyNP[i] + "|";

                            }
                            string[] disfilenames = text.Split('|');
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar1, label9);

                            //smallfile
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            text = "";
                            string[] tempStrings = { "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc" };
                            for (int i = 0; i < filenamessimplysamll.Length; i++)
                            {

                                if (tempStrings.Contains(filenamessimplysamll[i]))
                                {
                                    text = text + path1 + filenamessimplysamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path1 + filenamessimplysamll[i] + "_EN.xlsx" + "|";
                                    text = text + path1 + filenamessimplysamll[i] + "_FR.xlsx" + "|";
                                    text = text + path1 + filenamessimplysamll[i] + "_GER.xlsx" + "|";
                                }
                            }
                            filenames = text.Substring(0, text.Length - 1).Split('|');
                            text = "";
                            for (int i = 0; i < filenamessimplysamll.Length; i++)
                            {
                                if (tempStrings.Contains(filenamessimplysamll[i]))
                                {
                                    text = text + path2 + filenamessimplysamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path2 + filenamessimplysamll[i] + "_EN.xlsx" + "|";
                                    text = text + path2 + filenamessimplysamll[i] + "_FR.xlsx" + "|";
                                    text = text + path2 + filenamessimplysamll[i] + "_GER.xlsx" + "|";
                                }
                            
                            }
                            disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            CopyWithProgress(filenames, disfilenames, progressBar2, label6);

                            //cell index
                            text = "";
                            path1 = @"d:\ptw\lockfile\";
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            for (int i = 0; i < filenamessimplycell.Length; i++)
                            {
                                text = text + path1 + filenamessimplycell[i] + "|";

                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamessimplycell.Length; i++)
                            {
                                text = text + path2 + filenamessimplycell[i] + "|";

                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar3, label7);

                            //index
                            text = "";
                            path1 = @"d:\ptw\index\PrefaceNP\";
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            for (int i = 0; i < filenamessimplyindex.Length; i++)
                            {
                                text = text + path1 + filenamessimplyindex[i] + "|";
                                text = text + path1 + filenamessimplyindex[i] + "|";
                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamessimplyindex.Length; i++)
                            {
                                text = text + path2 + filenamessimplyindex[i] + "|";
                                text = text + path3 + filenamessimplyindex[i] + "|";
                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar4, label8);

                        }
                        if (flag == 2)
                        {
                            //prefaceNP
                            string text = "";

                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            for (int i = 0; i < filenamesnormarNP.Length; i++)
                            {
                                text = text + path1 + filenamesnormarNP[i] + "|";

                            }
                            string[] filenames = text.Split('|');
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }

                            text = "";
                            for (int i = 0; i < filenamesnormarNP.Length; i++)
                            {
                                text = text + path2 + filenamesnormarNP[i] + "|";

                            }
                            string[] disfilenames = text.Split('|');
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar1, label9);

                            //smallfile
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            text = "";
                            string[] tempStrings = { "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc" };
                            for (int i = 0; i < filenamesnormarsamll.Length; i++)
                            {
                              
                                if (tempStrings.Contains(filenamesnormarsamll[i]))
                                {
                                    text = text + path1 + filenamesnormarsamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path1 + filenamesnormarsamll[i] + "_EN.xlsx" + "|";
                                    text = text + path1 + filenamesnormarsamll[i] + "_FR.xlsx" + "|";
                                    text = text + path1 + filenamesnormarsamll[i] + "_GER.xlsx" + "|";
                                }
                            }
                            filenames = text.Substring(0, text.Length - 1).Split('|');
                            text = "";
                            for (int i = 0; i < filenamesnormarsamll.Length; i++)
                            {
                              
                                if (tempStrings.Contains(filenamesnormarsamll[i]))
                                {
                                    text = text + path2 + filenamesnormarsamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path2 + filenamesnormarsamll[i] + "_EN.xlsx" + "|";
                                    text = text + path2 + filenamesnormarsamll[i] + "_FR.xlsx" + "|";
                                    text = text + path2 + filenamesnormarsamll[i] + "_GER.xlsx" + "|";
                                }
                            }
                            disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            CopyWithProgress(filenames, disfilenames, progressBar2, label6);

                            //cell index
                            text = "";
                            path1 = @"d:\ptw\lockfile\";
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            for (int i = 0; i < filenamesnormarcell.Length; i++)
                            {
                                text = text + path1 + filenamesnormarcell[i] + "|";

                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamesnormarcell.Length; i++)
                            {
                                text = text + path2 + filenamesnormarcell[i] + "|";

                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar3, label7);

                            //index
                            text = "";
                            path1 = @"d:\ptw\index\PrefaceNP\";
                            string pathforblock = @"d:\ptw\";//Jintao: add block.index file to update 22042016
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";

                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            for (int i = 0; i < filenamesnormarlindex.Length; i++)
                            {
                                if (filenamesnormarlindex[i].IndexOf("block.index") != -1)
                                {
                                    text = text + pathforblock + filenamesnormarlindex[i] + "|";
                                    text = text + pathforblock + filenamesnormarlindex[i] + "|";
                                }
                                else
                                {
                                    text = text + path1 + filenamesnormarlindex[i] + "|";
                                    text = text + path1 + filenamesnormarlindex[i] + "|";
                                }

                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamesnormarlindex.Length; i++)
                            {
                                text = text + path2 + filenamesnormarlindex[i] + "|";
                                text = text + path3 + filenamesnormarlindex[i] + "|";
                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar4, label8);
                        }
                        if (flag == 3)
                        {
                            //prefaceNP
                            string text = "";

                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";

                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\";

                            }
                            for (int i = 0; i < filenamebothNP.Length; i++)
                            {
                                text = text + path1 + filenamebothNP[i] + "|";

                            }
                            string[] filenames = text.Split('|');
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }

                            text = "";
                            for (int i = 0; i < filenamebothNP.Length; i++)
                            {
                                text = text + path2 + filenamebothNP[i] + "|";

                            }
                            string[] disfilenames = text.Split('|');
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar1, label9);

                            //smallfile
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";

                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\divi\";
                            }
                            text = "";
                            string[] tempStrings = { "CmpcWacc", "CalculFCF", "DiscountedFCF", "MéthodesMixtes", "Multiples", "TransactionsComparables", "AutresCapitalisations", "GordonShapiroBates", "Goodwill", "PatrimonialAncAncc" };
                            for (int i = 0; i < filenamebothsamll.Length; i++)
                            {

                                if (tempStrings.Contains(filenamebothsamll[i]))
                                {
                                    text = text + path1 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path1 + filenamebothsamll[i] + "_EN.xlsx" + "|";
                                    text = text + path1 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                                    text = text + path1 + filenamebothsamll[i] + "_GER.xlsx" + "|";
                                }
                            }
                            filenames = text.Substring(0, text.Length - 1).Split('|');
                            text = "";
                            for (int i = 0; i < filenamebothsamll.Length; i++)
                            {

                                if (tempStrings.Contains(filenamebothsamll[i]))
                                {
                                    text = text + path2 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                                }
                                else
                                {
                                    text = text + path2 + filenamebothsamll[i] + "_EN.xlsx" + "|";
                                    text = text + path2 + filenamebothsamll[i] + "_FR.xlsx" + "|";
                                    text = text + path2 + filenamebothsamll[i] + "_GER.xlsx" + "|";
                                }
                            }
                            disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            CopyWithProgress(filenames, disfilenames, progressBar2, label6);

                            //cell index
                            text = "";
                            path1 = @"d:\ptw\lockfile\";
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";

                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            for (int i = 0; i < filenamebothcell.Length; i++)
                            {
                                text = text + path1 + filenamebothcell[i] + "|";

                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamebothcell.Length; i++)
                            {
                                text = text + path2 + filenamebothcell[i] + "|";

                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar3, label7);

                            //index
                            text = "";
                            path1 = @"d:\ptw\index\PrefaceNP\";
                            if (radioButton2.Checked)
                            {
                                path2 = @"\\PATRICK-6800\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICK-6800\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";
                            }
                            else if (radioButton1.Checked)
                            {
                                path2 = @"\\Dell-490\d\Solutions NOTA-PME\\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                            }
                            else if (radioButton3.Checked)
                            {
                                path2 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\PATRICKSTUDIO17\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";

                            }
                            else if (radioButton4.Checked)
                            {
                                path2 = @"\\Francis-7500\d\Solutions NOTA-PME\PilotageExcel_Bis\trunk\PilotageExcel\data\";
                                path3 = @"\\Francis-7500\d\Solutions NOTA-PME\NOTA-PME\Spreadsheets\index\PrefaceNP\";

                            }
                            for (int i = 0; i < filenamebothindex.Length; i++)
                            {
                                text = text + path1 + filenamebothindex[i] + "|";
                                text = text + path1 + filenamebothindex[i] + "|";
                            }
                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                filenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                filenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            text = "";
                            for (int i = 0; i < filenamebothindex.Length; i++)
                            {
                                text = text + path2 + filenamebothindex[i] + "|";
                                text = text + path3 + filenamebothindex[i] + "|";
                            }

                            if (text.Substring(0, text.Length - 1).Contains("|"))
                            {
                                disfilenames = text.Substring(0, text.Length - 1).Split('|');
                            }
                            else
                            {
                                disfilenames = new string[] { text.Substring(0, text.Length - 1) };
                            }
                            CopyWithProgress(filenames, disfilenames, progressBar4, label8);
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please choose one destination!");
                    }
                }
            }

        }
        private void CopyWithProgress(string[] filenames, string[] disfilenames,ProgressBar p1,Label lb)
        {
            // Display the ProgressBar control.
            p1.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            p1.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            p1.Maximum = filenames.Length;
            // Set the initial value of the ProgressBar.
            p1.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            p1.Step = 1;

            // Loop through all files to copy.
            for (int x = 1; x <= filenames.Length; x++)
            {
                // Copy the file and increment the ProgressBar if successful.
                try
                {
                    if (CopyFile(filenames[x - 1], disfilenames[x - 1]) == true)
                    {
                        // Perform the increment on the ProgressBar.
                        p1.PerformStep();
                    }
                }
                catch
                {
                    p1.PerformStep();
                }
            }
            p1.Visible = false;
            lb.Visible = true;
        }
        private bool CopyFile(string filepath,string destination)
        {
            bool flag = false;
            try
            {
                if (System.Environment.MachineName.ToUpper() == "PATRICK-6800")
                {
                    if (filepath.Contains("PATRICK-6800"))
                    {
                        filepath = filepath.Replace("\\\\PATRICK-6800\\d\\", "d:\\");
                        filepath = filepath.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                        filepath = filepath.Replace(@"\\\\PATRICK-6800\\d\\", "d:\\");
                        filepath = filepath.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                    }
                    if (destination.Contains("PATRICK-6800"))
                    {
                        destination = destination.Replace("\\\\PATRICK-6800\\d\\", "d:\\");
                        destination = destination.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                        destination = destination.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                        destination = destination.Replace(@"\\PATRICK-6800\\d\\", "d:\\");
                    }
                }
                if (File.Exists(filepath))
                {
                    File.Copy(filepath, destination, true);
                    flag = true;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR:" + ex.ToString());
            }
            return flag;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                backup("2");
                MessageBox.Show("Roll back finished");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (DateTime.Parse("2015/06/22") > DateTime.Parse("01/01/2014"))
            {
               // string x = DateTime.Parse("2015/06/22").ToString();
            }
            else
            {
                //string x = "";
            }
        }
    }
}
