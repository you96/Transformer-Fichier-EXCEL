using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using TransformEXCEL;

namespace WindowsExcelVSOT
{
    public partial class SubForm : Form
    {
        int number = 0;
        string filepath = @"\\PATRICK-6800\d\ptw\style nota-pme.xlsx";
        Form1 form1=null;
        public SubForm(int numb,Form1 form1)
        {
            this.form1 = form1;
            number = numb;
            InitializeComponent();
            string item = "";
            switch (number)
            {
                case 1: item = "PATRICK-6800"; comboBox1.Items.Add(item); filepath = @"\\PATRICK-6800\d\ptw\style nota-pme.xlsx"; break;
                case 2: item = "Dell-490"; comboBox1.Items.Add(item); filepath = @"\\Dell-490\d\ptw\style nota-pme.xlsx"; break;
                case 3: item = "DELL-E6000"; comboBox1.Items.Add(item); filepath = @"\\DELL-E6000\d\ptw\style nota-pme.xlsx"; break;
                default: break;
            }
            comboBox1.SelectedIndex = 0;
            switch (number)
            {
                case 1: item = "Tout"; comboBox2.Items.Add(item); item = "PATRICK-6800"; comboBox2.Items.Add(item); item = "DELL-E6000"; comboBox2.Items.Add(item); break;
                case 2: item = "Tout"; comboBox2.Items.Add(item); item = "DELL-E6000"; comboBox2.Items.Add(item); item = "DELL-490"; comboBox2.Items.Add(item); break;
                case 3: item = "Tout"; comboBox2.Items.Add(item); item = "PATRICK-6800"; comboBox2.Items.Add(item); item = "DELL-490"; comboBox2.Items.Add(item); break;
                default: break;
            }
            comboBox2.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (comboBox2.Text == "Tout")
            {
                switch (number)
                {
                    case 2: CopyFile(filepath, @"\\PATRICK-6800\d\ptw\style nota-pme.xlsx"); CopyFile(filepath, @"\\DELL-E6000\d\ptw\style nota-pme.xlsx"); break;
                    case 1: CopyFile(filepath, @"\\DELL-E6000\d\ptw\style nota-pme.xlsx"); CopyFile(filepath, @"\\DELL-490\d\ptw\style nota-pme.xlsx"); break;
                    case 3: CopyFile(filepath, @"\\DELL-490\d\ptw\style nota-pme.xlsx"); CopyFile(filepath, @"\\PATRICK-6800\d\ptw\style nota-pme.xlsx"); break;
                    default: break;
                }
            }
            else if (comboBox2.Text == "PATRICK-6800")
            {
                CopyFile(filepath, @"\\PATRICK-6800\d\ptw\style nota-pme.xlsx");
            }
            else if (comboBox2.Text == "DELL-E6000")
            {
                CopyFile(filepath, @"\\DELL-E6000\d\ptw\style nota-pme.xlsx");
            }
            else if (comboBox2.Text == "DELL-490")
            {
                CopyFile(filepath, @"\\DELL-490\d\ptw\style nota-pme.xlsx");
            }
            else
            {
                MessageBox.Show("select at least one destination");
            }
            if (form1 != null)
            {
                form1.setDateofStyle();
            }
            this.Close();
        }
        private bool CopyFile(string filepath, string destination)
        {
            bool flag = false;
            try
            {
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
            this.Close();
        }
    }
}
