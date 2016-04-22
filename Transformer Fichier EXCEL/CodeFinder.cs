using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace TransformEXCEL
{
    class CodeFinder
    {

        private String bookName;

        public String BookName
        {
            get { return bookName; }
            set { bookName = value; }
        }


        private String sheetName;

        
        public String SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }


        private Excel.Application xlsApp;

        public Excel.Application XlsApp
        {
            get { return xlsApp; }
            set { xlsApp = value; }
        }


        private  Excel.Workbook xlsWorkBook;

        public Excel.Workbook XlsWorkBook
        {
          get { return xlsWorkBook; }
          set { xlsWorkBook = value; }
        }

        private Excel.Worksheet xlsWorkSheet;

        public Excel.Worksheet XlsWorkSheet
        {
          get { return xlsWorkSheet; }
          set { xlsWorkSheet = value; }
        }

        /*
         * Constructeur de la classe;
         * Paramètres d'entrée:
         * Fonctionnement:
         * Erreurs Possibles:
         */
        //public CodeFinder(String bookName, String sheetName, Excel.Application xlsApp,Excel.Workbook xlsWorkBook, Excel.Worksheet xlsWorkSheet)
        //{
        //    this.BookName = bookName;
        //    this.SheetName = sheetName;
        //    this.XlsApp = XlsApp;
        //    this.XlsWorkBook = xlsWorkBook;
        //    this.XlsWorkSheet = xlsWorkSheet;
        //}

        public CodeFinder(Excel.Workbook xlWorkBook, Excel.Worksheet xlworksheet)
        {
            this.XlsWorkBook = xlWorkBook;
            this.XlsWorkSheet = xlworksheet;
        }

        private int compareString(String one, String two)
        {
            int result = 0;
            String[] s1 = new String[2];
            String[] s2 = new String[2];
            long i1;
            long i2;
            if (one == "")
            {
                one = "0";
            }
            if (two == "")
            {
                two = "0";
            }
            if (one.Contains('-') == true)
            {
                s1 = one.Split('-');
                if (two.Contains('-'))
                {
                    s2 = two.Split('-');
                    i1 = Int32.Parse(s1[0]);
                    i2 = Int32.Parse(s2[0]);
                    if (i1 == i2)
                    {
                        if (Int32.Parse(s1[1]) > Int32.Parse(s2[1]))
                        {
                            result = 1;
                        }
                        else if (Int32.Parse(s1[1]) < Int32.Parse(s2[1]))
                        {
                            result = 2;
                        }
                        else
                        {
                            result = 0;
                        }
                    }
                    else if (i1 > i2)
                    {
                        result = 1;
                    }
                    else { result = 2; }
                }
                else
                {
                    i1 = Int32.Parse(s1[0]);
                    i2 = Int32.Parse(two);
                    if (i1 == i2)
                    {
                        result = 1;

                    }
                    else if (i1 > i2)
                    {
                        result = 1;
                    }
                    else { result = 2; }
                }
            }
            else
                if (two.Contains('-') == true)
                {
                    s2 = two.Split('-');
                    if (one.Contains('-'))
                    {
                        s1 = one.Split('-');
                        i1 = Int32.Parse(s1[0]);
                        i2 = Int32.Parse(s2[0]);
                        if (i1 == i2)
                        {
                            if (Int32.Parse(s1[1]) > Int32.Parse(s2[1]))
                            {
                                result = 1;
                            }
                            else if (Int32.Parse(s1[1]) < Int32.Parse(s2[1]))
                            {
                                result = 2;
                            }
                            else
                            {
                                result = 0;
                            }
                        }
                        else if (i1 > i2)
                        {
                            result = 1;
                        }
                        else { result = 2; }
                    }
                    else
                    {
                        i1 = Int32.Parse(one);
                        i2 = Int32.Parse(s2[0]);
                        if (i1 == i2)
                        {
                            result = 2;

                        }
                        else if (i1 > i2)
                        {
                            result = 1;
                        }
                        else { result = 2; }
                    }
                }
                else
                {
                    i1 = (long)Int64.Parse(one);
                    i2 = (long)Int64.Parse(two);

                    if (i1 == i2)
                    {
                        result = 0;
                    }
                    else if (i1 > i2)
                    {
                        result = 1;
                    }
                    else
                    {
                        result = 2;
                    }
                }
            return result;
        }

        public int FindCodedColumn(String code, Excel.Range range)
        {
            //String cs = "Historique";
            /*
             * Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            rCnt = range.Rows.Count;
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "8000"))
                {
                    col = cCnt;
                    break;
                }
            }
             */
            int min;
            int max;
            int solution = 0;
            int res;
            int rCnt = range.Rows.Count;
            String s = "";
            min = 1;

            //l.Remove(targetSheet);
            object[,] values = (object[,])range.Value2;
            max = range.Columns.Count;
            while (min < max)
            {

                solution = min + (max - min) / 2;
                s = Convert.ToString(values[rCnt, solution]);


                res = compareString(code, s);
                if (res == 0)
                {
                    break;
                }
                if (res == 1)
                {
                    min = solution + 1;
                }
                else
                { max = solution; }
            }
           if (s.CompareTo(code) == 0)
            {
                return solution;
            }
           else
           {
               //MessageBox.Show("erreur! exception: colonne ne trouve pas: " + solution);
               return -1;
           }

        }

        public int FindCodedRow(String code, Excel.Range range)
        {
            //String cs = "Historique";
            /*
             * Excel.Range range = xlWorkSheet.UsedRange;
            object[,] values = (object[,])range.Value2;

            int rCnt = 0;
            int cCnt = 0;
            int col = 0;
            rCnt = range.Rows.Count;
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                string valuecellabs = Convert.ToString(values[rCnt, cCnt]);
                if (Regex.Equals(valuecellabs, "8000"))
                {
                    col = cCnt;
                    break;
                }
            }
             */
            int min;
            int max;
            int solution = 0;
            int res;
            int lCnt = range.Columns.Count;
            String s = "";
            min = 1;

            //l.Remove(targetSheet);
            object[,] values = (object[,])range.Value2;
            max = range.Rows.Count;
            while (min < max)
            {

                solution = min + (max - min) / 2;
                s = Convert.ToString(values[solution, lCnt]);


                res = compareString(code, s);
                if (res == 0)
                {
                    break;
                }
                if (res == 1)
                {
                    min = solution + 1;
                }
                else
                { max = solution; }
            }
            if (s.CompareTo(code) == 0)
            {
                return solution;
            }
            else
            {
                //MessageBox.Show("erreur! exception: ligne ne trouve pas: " + solution);
                return -1;
            }
        }

        public string FindCodedColumnHeader(String code, Excel.Range range)
        {
            int min;
            int max;
            int solution = 0;
            int res;
            int rCnt = range.Rows.Count;
            String s = "";
            min = 1;

            //l.Remove(targetSheet);
            object[,] values = (object[,])range.Value2;
            max = range.Columns.Count;
            while (min < max)
            {
                solution = min + (max - min) / 2;
                s = Convert.ToString(values[rCnt, solution]);

                res = compareString(code, s);
                if (res == 0)
                {
                    break;
                }
                if (res == 1)
                {
                    min = solution + 1;
                }
                else
                { max = solution; }
            }
            if (s.CompareTo(code) == 0)
            {
                return Number2String(solution, true);
                //return range.Columns[solution].ToString();
                //return solution;
            }
            else
            {
                //MessageBox.Show("erreur! exception: colonne ne trouve pas: " + solution);
                return "AAAA";
            }
        }

        private String Number2String(int number, bool isCaps)
        {
            int number1 = number / 27;
            int number2 = number - (number1 * 26);
            if (number2 > 26)
            {
                number1 = number1 + 1;
                number2 = number - (number1 * 26);
            }
            Char a = (Char)((isCaps ? 65 : 97) + (number1 - 1));
            Char b = (Char)((isCaps ? 65 : 97) + (number2 - 1));
            Char c = (Char)((isCaps ? 65 : 97) + (number - 1));
            string d = String.Concat(a, b);
            if (number <= 26)
                return c.ToString();
            else
                return d;
        }

        public void FlushAll()
        {
            object misValue = System.Reflection.Missing.Value;
            this.XlsApp.DisplayAlerts = true;
            this.XlsWorkBook.Close(true, misValue, misValue);
            this.XlsApp.Quit();
        }
    }
}
