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
using iTextSharp.text.pdf;

namespace Task_2
{
    internal class WorkFile
    {
        public static void CheckFile(string loadFile, string saveFile)
        {
            string nameFile = "\\Опись.docx";

            //Проверка на файл Опись
            if (!File.Exists(loadFile + nameFile))
            {
                Message.MessageError("В указанной папке не существует файла Опись");
                return;
            }

            int number = 0;
            int count = 0;

            Dictionary<string, bool> listFiles = new Dictionary<string, bool>();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile);

            //Фиксирования какие файлы и каталоги находятся в изначальном каталоге
            foreach (string str in mas) 
            {
                if (str.Contains(".pdf") || str.Contains(".PDF")|| Directory.Exists(str))
                    listFiles.Add(str.Replace(loadFile + "\\", ""), false);
            }

            //Перенос файла Опись в конечную папку
            File.Copy(loadFile + nameFile, saveFile + nameFile, true);

            //Открытие файла Опись
            var word = new Microsoft.Office.Interop.Word.Application();
            var doc = word.Documents.Open(saveFile + nameFile);

            var table = doc.Tables[1];

            string missingFiles = "";

            //Чтение таблицы из файла Опись
            for (int i = 3; i <= table.Rows.Count; i++)
            {
                string files = table.Rows[i].Cells[2].Range.Text.Replace("\a","").Replace("\t","").Replace("\n","").Replace("\r", "");

                int oldNumber = number;
                int oldCount = count;

                //Проверка на наличие документа или каталога из файла Опись в начальной папке
                if(listFiles.ContainsKey(files)) 
                {
                    InventoryDirectory(loadFile, saveFile, files, ref number, ref count);
                    listFiles[files] = true;
                }
                else
                {
                    if (listFiles.ContainsKey(files + ".pdf"))
                    {
                        InventoryFile(loadFile, saveFile, files, ref number, ref count);
                        listFiles[files + ".pdf"] = true;
                    }
                    else
                    {
                        missingFiles += files + "\n";
                    }
                }

                //Фиксация номера обработанного документа или каталогов и количество его страниц
                if(oldNumber != number) 
                {
                    if(oldNumber + 1 == number) 
                    {
                        table.Rows[i].Cells[1].Range.Text = number.ToString();
                    }
                    else
                    {
                        table.Rows[i].Cells[1].Range.Text = (oldNumber+1).ToString() + "-" + number.ToString();
                    }

                    if(oldCount + 1 == count) 
                    {
                        table.Rows[i].Cells[3].Range.Text = count.ToString();
                    }
                    else
                    {
                        table.Rows[i].Cells[3].Range.Text = (oldCount + 1).ToString() + "-" + count.ToString();
                    }
                }
            }

            //Учет лишних и отсутствующих документов
            CheckMissingFiles(saveFile, missingFiles);
            ExtraElement(loadFile, saveFile, listFiles);

            doc.Close();
            word.Quit();
        }



        //Рекурсивный метод для чтения каталогов
        private static void InventoryDirectory(string loadFile, string saveFile, string nameDirectory, ref int number, ref int count)
        {
            string name = (number+1).ToString();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile + "\\" + nameDirectory);

            Directory.CreateDirectory(saveFile + "\\" + nameDirectory);

            for (int i = 0; i < mas.Length; i++) 
            {
                if (File.Exists(mas[i])) 
                {
                    InventoryFile(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(saveFile + "\\", ""), ref number, ref count);
                }
                else
                {
                    InventoryDirectory(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(saveFile+"\\",""), ref number, ref count);
                }
            }

            name += "-" + (number).ToString();

            Directory.Move(saveFile + "\\" + nameDirectory, saveFile + "\\" + name + nameDirectory);
        }

        //Фиксация файла pdf
        private static void InventoryFile(string loadFile, string saveFile, string nameFile, ref int number, ref int count)
        {
            number++;

            //Копирование pdf файла из изначальной папки в конечную
            File.Copy(loadFile + "\\" + nameFile + ".pdf", saveFile + "\\" + number.ToString() +". "+ nameFile + ".pdf", true);

            //Получения количества страниц в pdf файле
            PdfReader pdf = new PdfReader(saveFile + "\\" + number.ToString() + ". " + nameFile + ".pdf");
            count += pdf.NumberOfPages;
            pdf.Close();
        }

        //Сохранение записи о не хватающих файлов
        private static void CheckMissingFiles(string saveFiles, string missingFiles)
        {

        }

        //Сохранения лишних файлов или файлов с неправильным названием
        private static void ExtraElement(string loadFile, string saveFile, Dictionary<string, bool> listFiles)
        {
            saveFile += "\\" + "Неопределенные";

            Directory.CreateDirectory(saveFile);

            foreach (string str in listFiles.Keys)
            {
                if (!listFiles[str]) 
                {
                    if(Directory.Exists(loadFile +"\\"+str))
                    {
                        ExtraDirectory(loadFile, saveFile, str);
                    }
                    else
                    {
                        if (listFiles.ContainsKey(str))
                        {
                            ExtraFile(loadFile, saveFile, str);
                        }
                    }
                }
            }
        }

        private static void ExtraDirectory(string loadFile, string saveFile, string nameDirectory)
        {

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile + "\\" + nameDirectory);

            Directory.CreateDirectory(saveFile + "\\" + nameDirectory);

            for (int i = 0; i < mas.Length; i++)
            {
                if (File.Exists(mas[i]))
                {
                    ExtraFile(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(saveFile + "\\", ""));
                }
                else
                {
                    ExtraDirectory(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(saveFile + "\\", ""));
                }
            }

            Directory.Move(saveFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory);
        }

        private static void ExtraFile(string loadFile, string saveFile, string nameFile)
        {
            //Копирование pdf файла из изначальной папки в конечную
            File.Copy(loadFile + "\\" + nameFile, saveFile + "\\" + nameFile, true);
        }
    }
}
