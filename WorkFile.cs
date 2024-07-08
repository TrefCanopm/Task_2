using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using System.Globalization;

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

            Microsoft.Office.Interop.Word.Application app = null;
            Document word = null;

            Dictionary<string, bool> listFiles = new Dictionary<string, bool>();

            try
            { 
            try
            {
                string[] mas = System.IO.Directory.GetFileSystemEntries(loadFile);

                //Фиксирования какие файлы и каталоги находятся в изначальном каталоге
                foreach (string str in mas)
                {
                    if (str.Contains(".pdf") || str.Contains(".PDF") || Directory.Exists(str))
                        listFiles.Add(str.Replace(loadFile + "\\", ""), false);
                }
            }
            catch (Exception e)
            {
                Message.MessageError("Ошибка чтения файлов из изначальной папки");
            }

            try
            {
                //Перенос файла Опись в конечную папку
                File.Copy(loadFile + nameFile, saveFile + nameFile, true);
            }
            catch
            {
                Message.MessageError("Ошибка копирования файла Опись");
            }

            //Открытие файла Опись

            try
            {
                app = new Microsoft.Office.Interop.Word.Application();
                word = app.Documents.Open(saveFile + nameFile);
            }
            catch
            {
                Message.MessageError("Ошибка открытия с копированного файла");
                if (word != null)
                   word.Close();
               if (app != null)
                    app.Quit();
            }

            List<string> missingFiles = new List<string>();
            try
            {
                var table = word.Tables[1];

                missingFiles = new List<string>();

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
                            if (listFiles.ContainsKey(files + ".PDF"))
                            {
                                table.Rows[i].Cells[2].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Color.Green);
                                InventoryFile(loadFile, saveFile, files + ".PDF", ref number, ref count);
                                listFiles[files + ".PDF"] = true;
                            }
                            else
                            {
                                table.Rows[i].Cells[2].Range.Font.Color = (WdColor)ColorTranslator.ToOle(Color.Red);
                                missingFiles.Add(files);
                            }
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
            }
            catch(Exception e)
            {

                Message.MessageError("Не отслеживаемая ошибка"+"\n"+e.ToString());
            }
            //Указание даты формирования отчета в верхний колонтикул

            try
            {
                foreach (Section section in word.Sections)
                {
                    var headers = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headers.Text = "Дата создания отчета " + DateTime.Now.ToString("d", CultureInfo.CreateSpecificCulture("de-DE"));
                }
            }
            catch
            {
                Message.MessageError("Ошибка указания даты в файл Опись");
            }
            //Строка сводки по документам
            word.Words.Last.InsertBefore(" Документов " + number + "\n Количество листов " + count + "\n");



            //Учет лишних и отсутствующих документов
            CheckMissingFiles(saveFile, missingFiles, word);
            ExtraElement(loadFile, saveFile, listFiles, app, word);

            word.Close();
            app.Quit();

            Message.MessageNotification("Опись произведена");
            }
            catch (Exception ex)
            {
                Message.MessageError(ex.ToString());
                if (word != null)
                    word.Close();
                if (app != null)
                    app.Quit();
            }
        }



        //Рекурсивный метод для чтения каталогов
        private static void InventoryDirectory(string loadFile, string saveFile, string nameDirectory, ref int number, ref int count)
        {
            string name = (number + 1).ToString();
            string[] mas = new string[0];
            try
            {
                mas = System.IO.Directory.GetFileSystemEntries(loadFile + "\\" + nameDirectory);
            }
            catch
            {
                Message.MessageError("Ошибка чтения каталогов из основной папки");
            }

            try
            {
                Directory.CreateDirectory(saveFile + "\\" + nameDirectory);
            }
            catch
            {
                Message.MessageError("Ошибка создания каталога");
            }

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

            try
            {
                Directory.Move(saveFile + "\\" + nameDirectory, saveFile + "\\" + name + nameDirectory);
            }
            catch
            {
                Message.MessageError("Ошибка изменения названия необходимых скопированных каталогов");
            }
        }

        //Фиксация файла pdf
        private static void InventoryFile(string loadFile, string saveFile, string nameFile, ref int number, ref int count)
        {
            number++;

            //Копирование pdf файла из изначальной папки в конечную
            try
            {
                File.Copy(loadFile + "\\" + nameFile, saveFile + "\\" + number.ToString() + ". " + nameFile, true);
            }
            catch
            {
                Message.MessageError("Ошибка копирования элементов необходимых файлов");
            }

            //Получения количества страниц в pdf файле

            try
            {
                PdfReader pdf = new PdfReader(saveFile + "\\" + number.ToString() + ". " + nameFile);
                count += pdf.NumberOfPages;
                pdf.Close();
            }
            catch
            {
                Message.MessageError("Ошибка чтения количества страниц из pdf файла");
            }

        }

        //Сохранение записи о не хватающих файлов
        private static void CheckMissingFiles(string saveFiles, List<string> missingFiles, Document doc)
        {
            int i = 1;

            try
            {
                doc.Words.Last.InsertBefore("Не найденные файлы:" + "\n");
                foreach (string str in missingFiles)
                {
                    doc.Words.Last.InsertBefore(i + ") " + str + "\n");
                    i++;
                }
            }
            catch
            {
                Message.MessageError("Ошибка фиксации не найденных файлов в документ Опись");
            }
        }

        //Сохранения лишних файлов или файлов с неправильным названием
        private static void ExtraElement(string loadFile, string saveFile, Dictionary<string, bool> listFiles, Microsoft.Office.Interop.Word.Application app, Document word)
        {
            try
            {
                saveFile += "\\" + "Неопределенные";
                Directory.CreateDirectory(saveFile);
            }
            catch
            {
                Message.MessageError("Ошибка создания каталога с неопределенными файлами");
            }

            word.Words.Last.InsertBefore("\nНеопределенные файлы:\n");

            object number = 1;
            ListTemplate template = app.ListGalleries[WdListGalleryType.wdNumberGallery].ListTemplates.get_Item(ref number);

            Paragraph paragraph = null;
            var range = word.Range();

            int listLevel = 1;

            foreach (string str in listFiles.Keys)
            {
                if (!listFiles[str])
                {
                    paragraph = range.Paragraphs.Add();
                    paragraph.Range.Text = str.Replace(".pdf", "").Replace("PDF", "");

                    paragraph.Range.SetListLevel((short)listLevel);
                    paragraph.Range.ListFormat.ApplyListTemplateWithLevel(template, ContinuePreviousList: true, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior, ApplyLevel: 1);
                    paragraph.Range.InsertParagraphAfter();

                    if (Directory.Exists(loadFile + "\\" + str))
                    {
                        ExtraDirectory(loadFile, saveFile, str, paragraph, range, template, listLevel + 1);
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

        private static void ExtraDirectory(string loadFile, string saveFile, string nameDirectory, Paragraph paragraph, Microsoft.Office.Interop.Word.Range range, ListTemplate template, int listLevel)
        {
            string[] mas = new string[0];

            try
            {
                mas = System.IO.Directory.GetFileSystemEntries(loadFile + "\\" + nameDirectory);
            }
            catch
            {
                Message.MessageError("Ошибка чтения файлов в неопределенном каталоге");
            }

            try
            {
                Directory.CreateDirectory(saveFile + "\\" + nameDirectory);
            }
            catch
            {
                Message.MessageError("Ошибка переноса неопределенного котолога");
            }

            for (int i = 0; i < mas.Length; i++)
            {
                paragraph = range.Paragraphs.Add();
                paragraph.Range.Text = mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", "").Replace(".pdf", "").Replace("PDF", "");

                paragraph.Range.SetListLevel((short)listLevel);
                paragraph.Range.ListFormat.ApplyListTemplateWithLevel(template, ContinuePreviousList: true, ApplyLevel: 1);
                paragraph.Range.InsertParagraphAfter();

                if (File.Exists(mas[i]))
                {
                    ExtraFile(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", ""));
                }
                else
                {
                    ExtraDirectory(loadFile + "\\" + nameDirectory, saveFile + "\\" + nameDirectory, mas[i].Replace(loadFile + "\\" + nameDirectory + "\\", ""), paragraph, range, template, listLevel + 1);
                }
            }
        }

        private static void ExtraFile(string loadFile, string saveFile, string nameFile)
        {
            try
            {
                //Копирование pdf файла из изначальной папки в конечную
                File.Copy(loadFile + "\\" + nameFile, saveFile + "\\" + nameFile, true);
            }
            catch
            {
                Message.MessageError("Ошибка копирования неопределенных файлов");
            }
        }
    }
}