using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Task_2
{
    internal class WorkFile
    {
        public static void ChekFile(string fileLoad, string fileSave)
        {
            if (!File.Exists(fileLoad + "\\Опись.docx"))
            {
                Message.MessageError("В указанной папке не существует файла Опись");
                return;
            }

            int number = 0;
            int count = 0;

            Dictionary<string, bool> files = new Dictionary<string, bool>();

            string str1 = "";

            string[] mas = System.IO.Directory.GetFileSystemEntries(fileLoad);

            foreach (string str in mas) 
            {
                files.Add(str.Replace(fileLoad + "\\", ""), false);
            }

            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = word.Documents.Open(fileLoad + "\\Опись.docx");

            var table = doc.Tables[0];

            for(int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                if (files.ContainsKey(row.Cells[1].Range.Text))
                {
                    str1 += row.Cells[1].Range.Text;
                }
            }

        }

        private static void InventoryDirectory(string fileLoad, string fileSave, string nameDirectory, ref int number, ref int count)
        {
            string name = (number+1).ToString();

            string[] mas = System.IO.Directory.GetFileSystemEntries(fileLoad);
        }

        private static void InventoryFile(string fileLoad, string fileSave, ref int number, ref int count)
        {

        }
    }
}
