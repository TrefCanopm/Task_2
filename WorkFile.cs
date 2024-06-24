﻿using System;
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

            Microsoft.Office.Interop.Word.Application word = null;
            Document doc = null;

            Dictionary<string, bool> listFiles = new Dictionary<string, bool>();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile);

            //Фиксирования какие файлы и каталоги находятся в изначальном каталоге
            foreach (string str in mas)
            {
                if (str.Contains(".pdf") || str.Contains(".PDF") || Directory.Exists(str))
                    listFiles.Add(str.Replace(loadFile + "\\", ""), false);
            }

            try
            {
                //Перенос файла Опись в конечную папку
                File.Copy(loadFile + nameFile, saveFile + nameFile, true);

                //Открытие файла Опись
                word = new Microsoft.Office.Interop.Word.Application();
                doc = word.Documents.Open(saveFile + nameFile);

                var table = doc.Tables[1];

                List<string> missingFiles = new List<string>();

                //Чтение таблицы из файла Опись
                for (int i = 3; i <= table.Rows.Count; i++)
                {
                    string files = table.Rows[i].Cells[2].Range.Text.Replace("\a", "").Replace("\t", "").Replace("\n", "").Replace("\r", "");

                    int oldNumber = number;
                    int oldCount = count;

                    //Проверка на наличие документа или каталога из файла Опись в начальной папке
                    if (listFiles.ContainsKey(files))
                    {
                        table.Rows[i].Cells[2].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Color.Green);
                        InventoryDirectory(loadFile, saveFile, files, ref number, ref count);
                        listFiles[files] = true;
                    }
                    else
                    {
                        if (listFiles.ContainsKey(files + ".pdf"))
                        {
                            table.Rows[i].Cells[2].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Color.Green);
                            InventoryFile(loadFile, saveFile, files + ".pdf", ref number, ref count);
                            listFiles[files + ".pdf"] = true;
                        }
                        else
                        {
                            table.Rows[i].Cells[2].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Color.Red);
                            missingFiles.Add(files);
                        }
                    }

                    //Фиксация номера обработанного документа или каталогов и количество его страниц
                    if (oldNumber != number)
                    {
                        if (oldNumber + 1 == number)
                        {
                            table.Rows[i].Cells[1].Range.Text = number.ToString();
                        }
                        else
                        {
                            table.Rows[i].Cells[1].Range.Text = (oldNumber + 1).ToString() + "-" + number.ToString();
                        }

                        if (oldCount + 1 == count)
                        {
                            table.Rows[i].Cells[3].Range.Text = count.ToString();
                        }
                        else
                        {
                            table.Rows[i].Cells[3].Range.Text = (oldCount + 1).ToString() + "-" + count.ToString();
                        }
                    }
                }

                //Указание даты формирования отчета в верхний колонтикул
                foreach (Section section in doc.Sections)
                {
                    var headers = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headers.Text = "Дата создания отчета " + DateTime.Now.ToString("dd.mm.yy");
                }

                //Строка сводки по документам
                doc.Words.Last.InsertBefore(" Документов " + number + "\n Количество листов " + count + "\n");



                //Учет лишних и отсутствующих документов
                CheckMissingFiles(saveFile, missingFiles, doc);
                ExtraElement(loadFile, saveFile, listFiles, doc);

                doc.Close();
                word.Quit();

                Message.MessageNotification("Опись произведена");
            }
            catch (Exception ex)
            {
                Message.MessageError(ex.ToString());
                if (doc != null)
                    doc.Close();
                if (word != null)
                    word.Quit();
            }
        }



        //Рекурсивный метод для чтения каталогов
        private static void InventoryDirectory(string loadFile, string saveFile, string nameDirectory, ref int number, ref int count)
        {
            string name = (number + 1).ToString();

            string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile + "\\" + nameDirectory);

            Directory.CreateDirectory(saveFile + "\\" + nameDirectory);

            for (int i = 0; i < mas.Length; i++)
            {
                if (File.Exists(mas[i]))
                {
                    InventoryFile(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", ""), ref number, ref count);
                }
                else
                {
                    InventoryDirectory(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\", ""), ref number, ref count);
                }
            }

            name += "-" + (number).ToString() + ". ";

            Directory.Move(saveFile + "\\" + nameDirectory, saveFile + "\\" + name + nameDirectory);
        }

        //Фиксация файла pdf
        private static void InventoryFile(string loadFile, string saveFile, string nameFile, ref int number, ref int count)
        {
            number++;

            //Копирование pdf файла из изначальной папки в конечную
            File.Copy(loadFile + "\\" + nameFile, saveFile + "\\" + number.ToString() + ". " + nameFile, true);

            //Получения количества страниц в pdf файле
            PdfReader pdf = new PdfReader(saveFile + "\\" + number.ToString() + ". " + nameFile);
            count += pdf.NumberOfPages;
            pdf.Close();
        }

        //Сохранение записи о не хватающих файлов
        private static void CheckMissingFiles(string saveFiles, List<string> missingFiles, Document doc)
        {
            int i = 1;

            doc.Words.Last.InsertBefore("Не найденные файлы:" + "\n");
            foreach (string str in missingFiles) 
            {
                doc.Words.Last.InsertBefore(i + ") " + str + "\n");
                i++;
            }
        }

        //Сохранения лишних файлов или файлов с неправильным названием
        private static void ExtraElement(string loadFile, string saveFile, Dictionary<string, bool> listFiles, Document doc)
        {
            saveFile += "\\" + "Неопределенные";

            Directory.CreateDirectory(saveFile);

            foreach (string str in listFiles.Keys)
            {
                if (!listFiles[str])
                {
                    if (Directory.Exists(loadFile + "\\" + str))
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
                    ExtraFile(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", ""));
                }
                else
                {
                    ExtraDirectory(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", ""));
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