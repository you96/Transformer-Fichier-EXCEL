using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace TransformEXCEL
{
    public class Index_Files
    {
        private string connexionString;
        private DataTable dataTable;
        private OleDbDataAdapter da;
        private String bookPath;
        private String divBookPath;

        public String DivBookPath
        {
            get { return divBookPath; }
            set { divBookPath = value; }
        }


        public String BookPath
        {
          get { return bookPath; }
          set { bookPath = value; }
        }

        private String bookName;

        public String BookName
        {
          get { return bookName; }
          set { bookName = value; }
        }

        private List<List<List<String>>> sheetsIndexList;

        public List<List<List<String>>> SheetsIndexList
        {
            get { return sheetsIndexList; }
            set { sheetsIndexList = value; }
        }
        private String indexDirectoryPath;

        public String IndexDirectoryPath
        {
            get { return indexDirectoryPath; }
            set { indexDirectoryPath = value; }
        }

    
        public Index_Files(String indexdirectoryPath, String bookName,String bookPath, String divBookPath)
        {
            this.BookName = bookName;
            this.BookPath = bookPath;
            this.DivBookPath = divBookPath;
            this.indexDirectoryPath = indexdirectoryPath;
            SheetsIndexList = new List<List<List<String>>>();
          // connexionString =  string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'", this.BookPath);
            if (this.BookPath.Contains(".xlsx") == true)// 2007 -2010
            {
                connexionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=NO';", this.BookPath);
            }
            else // 2003
            {
                connexionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'", this.BookPath);
            }
        }

        public void createColumnIndexFile(String filePath, String sql)
        {
            dataTable = new DataTable();
            da = new OleDbDataAdapter(sql, connexionString);
            da.Fill(dataTable);
            if (dataTable.Rows.Count != 0)//Si la table est vide
            {
            
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            
                for (int i = 0; i < dataTable.Columns.Count; ++i)
                {
                    sw.WriteLine((dataTable.Rows[dataTable.Rows.Count - 1][i].ToString()));
                }

                sw.Close();

                file.Close();
            }

        }
        public void createColumnIndexFile(String filePath, String sql,Excel.Workbook xlbook,bool state)
        {
            Excel.Worksheet xlsheet = xlbook.Sheets.get_Item(sql) as Excel.Worksheet;
            Excel.Range range = xlsheet.UsedRange;
            if (sql == "Historique" && state)
            {
                range = xlsheet.get_Range("A1", "BL661");
            }
           
            object[,] values = (object[,])range.Value2;
            int ccount = range.Columns.Count;
                FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(file);
                if (values != null)
                {
                    for (int i = 1; i <= ccount; i++)
                    {
                        if (values[range.Rows.Count, i] != null)
                        {
                            string sss = values[range.Rows.Count, i].ToString();
                            sw.WriteLine((values[range.Rows.Count, i].ToString()));
                        }
                        else
                        {
                            sw.WriteLine("");
                        }
                    }
                }
                sw.Close();

                file.Close();
            

        }
        public void createRowIndexFile(String filePath, String sql)
        {
            dataTable = new DataTable();
            da = new OleDbDataAdapter(sql, connexionString);
            da.Fill(dataTable);
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            if (dataTable.Rows.Count > 1334)
            {
                string ssss = dataTable.Rows[0][dataTable.Columns.Count - 1].ToString();
            }
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                string sss="";
                if (dataTable.Rows[i][dataTable.Columns.Count - 1].ToString() == "")
                {
                    if (dataTable.Columns.Count - 2>0)
                    {
                        if (dataTable.Rows[i][dataTable.Columns.Count - 2].ToString().Contains("-"))
                        {
                            if (dataTable.Rows[i][dataTable.Columns.Count - 2].ToString().IndexOf("-")!=0)
                            {
                                sss = Convert.ToInt32(dataTable.Rows[i][dataTable.Columns.Count - 2].ToString().Split('-')[0]) + "-" + Convert.ToInt32(dataTable.Rows[i][dataTable.Columns.Count - 2].ToString().Split('-')[1]);
                                if (Convert.ToInt32(dataTable.Rows[i][dataTable.Columns.Count - 2].ToString().Split('-')[1]) == 0)
                                {
                                    sss = sss.Split('-')[0];
                                }
                            }
                        }
                        else
                        {
                            sss = "";
                        }
                    }
                    else
                    {
                        sss = "";
                    }
                }
                else
                {
                    sss = dataTable.Rows[i][dataTable.Columns.Count - 1].ToString();
                }
                sw.WriteLine(sss);
            }

            sw.Close();

            file.Close();

        }
        public void createRowIndexFile(String filePath, String sql, Excel.Workbook xlbook,bool state)
        {
            Excel.Worksheet xlsheet = xlbook.Sheets.get_Item(sql) as Excel.Worksheet;
            Excel.Range range = xlsheet.UsedRange;
            if (sql == "Historique" && state)
            {
                range = xlsheet.get_Range("A1","BL661");
            }
            object[,] values = (object[,])range.Value2;
            int ccount = range.Columns.Count;
            FileStream file = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            if (values != null)
            {
                for (int i = 1; i <= range.Rows.Count; i++)
                {
                    if (values[i, ccount] != null)
                    {
                        sw.WriteLine(values[i, ccount].ToString());
                    }
                    else
                    {
                        sw.WriteLine("");
                    }
                }
            }
           
               
            
            

            sw.Close();

            file.Close();

        }
        public void createindexFiles(String Sheet, Excel.Workbook xlbook, bool state)
        {
            string sql = string.Format("Select * From [{0}]", Sheet + "$");
            String filePath1 = this.IndexDirectoryPath + "/" + this.BookName +"/" + Sheet + "_Columns.index";
            String filePath2 = this.IndexDirectoryPath + "/" + this.BookName + "/" + Sheet + "_Rows.index";
            if (Sheet == "Historique" && state)
            {
                filePath1 = this.IndexDirectoryPath + "/" + this.BookName + "/" + "Historique-s_Columns.index";
                filePath2 = this.IndexDirectoryPath + "/" + this.BookName + "/" + "Historique-s_Rows.index";
            }
            string tab1 = Sheet ;
            createColumnIndexFile(filePath1, tab1,xlbook,state);
           // createColumnIndexFile(filePath1, sql);
            //createRowIndexFile(filePath2, sql);
            createRowIndexFile(filePath2, tab1, xlbook, state);
           
        }

        public List<String> getSheetColumns(String sheet)
        {

            string f = IndexDirectoryPath + "/" + this.BookName + "/" + sheet + "_Columns.index";

            List<string> lines = new List<string>();
            lines.Add(sheet);

            using (
                
                StreamReader r = new StreamReader(f))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            return lines;
        }

        public List<String> getSheetRows(String Sheet)
        {
            string f = IndexDirectoryPath + "/" + this.BookName + "/" + Sheet + "_Rows.index";

            List<string> lines = new List<string>();
            lines.Add(Sheet);

            using (StreamReader r = new StreamReader(f))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            return lines;
        }
        public List<List<String>> ColsRowsSheet( String Sheet)
        {
            List<List<String>> l = new List<List<String>>();
            l.Add(getSheetColumns(Sheet));
            l.Add(getSheetRows(Sheet));
            return l;
        }


        private String[] GetExcelSheetNames(string targetBook)
        {
            OleDbConnection objConn = null;
            String s = "";
            dataTable = new DataTable();
        
            try
            {
                objConn = new OleDbConnection(connexionString);
                objConn.Open();

                dataTable = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dataTable == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dataTable.Rows.Count];
                int i = 0;
                int counter = 0;

                foreach (DataRow row in dataTable.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Contains("Comptes annuels"))
                    {
                        s = row["TABLE_NAME"].ToString();
                        s = row["TABLE_NAME"].ToString()[row["TABLE_NAME"].ToString().Length - 1].ToString();
                    }
                    if (row["TABLE_NAME"].ToString()[row["TABLE_NAME"].ToString().Length - 1].CompareTo('$') == 0 || row["TABLE_NAME"].ToString()[row["TABLE_NAME"].ToString().Length - 2].CompareTo('$') == 0)
                    {

                        if (!row["TABLE_NAME"].ToString().Contains("Don"))
                        {
                            s = row["TABLE_NAME"].ToString().Replace("'", "");

                            if (s.CompareTo("$") != 0)
                            {
                                excelSheets[i] = s;
                                counter++;
                            }
                            ++i;
                        }
                       

                    }
                }
                String[] res = new String[counter];
                for (int j = 0; j < counter; ++j)
                {
                    res[j] = excelSheets[j];
                }
                return res;

            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                // Clean up.

                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dataTable != null)
                {
                    dataTable.Dispose();
                }
            }
        }

        public void CreateDivFileSheets(String divFilePath)
        {
            OleDbConnection objConn = null;
            String connexionString_bis = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'", this.DivBookPath);
            FileStream file = new FileStream(divFilePath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(file);
            String s = "";
            dataTable = new DataTable();

            try
            {
                objConn = new OleDbConnection(connexionString_bis);
                objConn.Open();

                dataTable = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dataTable == null)
                {
                    return;
                }

                String[] excelSheets = new String[dataTable.Rows.Count];
                int i = 0;
                int counter = 0;

                foreach (DataRow row in dataTable.Rows)
                {

                    if (row["TABLE_NAME"].ToString()[row["TABLE_NAME"].ToString().Length - 1].CompareTo('$') == 0)
                    {
                        s = row["TABLE_NAME"].ToString();
                        if (s.CompareTo("$") != 0)
                        {
                            excelSheets[i] = s;
                            counter++;
                        }
                        ++i;
                    }
                }
                String[] res = new String[counter];
                for (int j = 0; j < counter; ++j)
                {
                    res[j] = excelSheets[j];
                    sw.WriteLine(res[j].ToString());
                }
                //return res;
                sw.Close();
                file.Close();

            }
            catch (Exception ex)
            {
                return;
            }
            finally
            {
                // Clean up.

                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dataTable != null)
                {
                    dataTable.Dispose();
                }
            }
        }


        public String ShowList(List<String> list)
        {
            int i = 0;
            String s = "";
            if (list != null)
            {
                while (i < list.Count)
                {
                    s = s + list[i] + "|";
                    i++;
                }
            }
            else { s = "NULL!!"; };
            return s;
        }

        public List<List<List<String>>> AddSheets()
        {
            String[] sheets = GetExcelSheetNames(this.BookName);
            String str;
            for (int i = 0; i < sheets.Length; ++i)
            {
                str = sheets[i].Replace("#", ".");
                AddSheet(SheetsIndexList, str.Remove(str.Length - 1, 1));
            }
            return SheetsIndexList;
        }
        public void CreateFiles(bool state)
        {

            String[] sheets = GetExcelSheetNames(this.BookName);

            String str;
            Directory.CreateDirectory(this.IndexDirectoryPath +"/" + BookName);
            String divfilePath = this.IndexDirectoryPath + "/" + this.BookName + "/" + this.BookName + ".index";
            CreateDivFileSheets(divfilePath);
            Excel.Application app = new Excel.ApplicationClass();
            app.Visible = true;
            app.DisplayAlerts = false;
            Excel.Workbook xlWorkBook;
            if (state)
            {
                xlWorkBook = app.Workbooks.Open("D:\\ptw\\prefaceNPS.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            }
            else
            {
                xlWorkBook = app.Workbooks.Open("D:\\ptw\\prefaceNP.xlsx", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            }
            for (int i = 0; i < sheets.Length; ++i)
            {
                str = sheets[i].Replace("#", ".");
                if (str.Remove(str.Length - 1, 1) == "Historique" || str.Remove(str.Length - 1, 1) == "Comptes annuels")
                {
                    createindexFiles(str.Remove(str.Length - 1, 1), xlWorkBook, state);
                }
            }
            xlWorkBook.Close();
            app.Quit();
        }

        public void AddSheet(List<List<List<String>>> infosheets,String Sheet)
        {
            infosheets.Add(ColsRowsSheet(Sheet));
        }


        public int ExistsSheet( String Sheet)
        {
            int result = -1;
            int i = 0;
            while (i < SheetsIndexList.Count)
            {
                if ( SheetsIndexList[i][0][0].CompareTo(Sheet) == 0)
                {
                    result = i;
                    break;
                }
                ++i;
            }
            return result;
        }
        public List<String> findSheetRows(String Sheet)
        {
            if (ExistsSheet(Sheet) != -1)
            {
                return SheetsIndexList[ExistsSheet(Sheet)][1];
            }
            else return null;
        }

        public List<String> findSheetColumns(String Sheet)
        {
            if (ExistsSheet(Sheet) != -1)
            {
                return SheetsIndexList[ExistsSheet(Sheet)][0];
            }
            else return null;
        }
    }
}