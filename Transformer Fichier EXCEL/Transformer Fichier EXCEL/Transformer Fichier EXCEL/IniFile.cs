﻿using System;
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

namespace INI
{
    class IniFile
    {
        public string filePath;
        
        //  declarer ini api 
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        ///  <summary>
        ///  construit fonction INIPath
        ///  </summary>
        ///  <param  name="INIPath">ini nom</param>  
        public IniFile(string INIPath)
        {
            filePath = INIPath;
        }
        ///  <summary>
        ///  ecrire ini
        ///  </summary>
        ///  <param  name="Section">Section</param>
        ///  <param  name="Key">Key</param>
        ///  <param  name="value">value</param>
        public void WriteInivalue(string Section, string Key, string value)
        {
            WritePrivateProfileString(Section, Key, value, this.filePath);
        }
        ///  <summary>
        ///  lire ini specifique part
        ///  </summary>
        ///  <param  name="Section">Section</param>
        ///  <param  name="Key">Key</param>
        ///  <returns>String</returns>  
        public string ReadInivalue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(1024);
            int i = GetPrivateProfileString(Section, Key, "erreur", temp, 1024, this.filePath);
            return temp.ToString();
        }
        //lire
    }
}
