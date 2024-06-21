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
using System.Reflection;

namespace Task_2
{
    internal class WorkFile
    {
        public static void CheckFile(string loadFile, string saveFile)
        {
            string nameFile = "\\Опись.docx";

            if (!File.Exists(loadFile + nameFile))
            {
                Message.MessageError("В указанной папке не существует файла Опись");
                return;
            }

            int number = 0;
            int count = 0;

            Dictionary<string, bool> listFiles = new Dictionary<string, bool>();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile);

            foreach (string str in mas) 
            {
                listFiles.Add(str.Replace(loadFile + "\\", ""), false);
            }

            File.Copy(loadFile + nameFile, saveFile + nameFile);

            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = word.Documents.Open(saveFile + nameFile);

            var table = doc.Tables[1];

            string missingFiles = "";

            for (int i = 2; i <= table.Rows.Count; i++)
            {
                string files = table.Rows[i].Cells[2].Range.Text;

                if(listFiles.ContainsKey(files)) 
                {
                    InventoryDirectory(loadFile, saveFile, files, ref number, ref count);
                }
                else
                {
                    if (listFiles.ContainsKey(files + ".pdf"))
                    {
                        InventoryFile(loadFile, saveFile, files, ref number, ref count);
                    }
                    else
                    {
                        missingFiles += files + "\n";
                    }
                }
            }

            ChekMissingFiles(saveFile, missingFiles);
            ExtraFiles(loadFile, saveFile, listFiles);

            doc.Close();
            word.Quit();
        }

        private static void InventoryDirectory(string loadFile, string saveFile, string nameDirectory, ref int number, ref int count)
        {
            string name = (number+1).ToString();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile);
        }

        private static void InventoryFile(string loadFile, string saveFile, string nameFile, ref int number, ref int count)
        {

        }

        private static void ChekMissingFiles(string saveFiles, string missingFiles)
        {

        }

        private static void ExtraFiles(string loadFile, string saveFile, Dictionary<string, bool> listFile)
        {

        }
    }
}
