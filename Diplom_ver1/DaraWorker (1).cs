using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Diplom_ver1
{

    internal class DataWorker
    {
        // создаем класс для работы с данными
        Data_Storage data = new Data_Storage();
        
        // открываем эксель таблицу
        public bool OpenXLSX(string funkName)
        {
            var choofdlog = new OpenFileDialog
            {
                Filter = "Excel Лист(xlsx)|*.xlsx",
                FilterIndex = 1,
                Multiselect = false
            };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                data.FileName = choofdlog.FileName; // путь к Excel файлу
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл");
                return true;
            }

            try
            {
                switch (funkName)
                {

                    case "Information":
                        data.Worksheet = new ClosedXML.Excel.XLWorkbook(data.FileName).Worksheets.First();
                        FromXLSL_toForm();
                        break;
                    case "Raiting":
                        data.WorksheetRaiting = new ClosedXML.Excel.XLWorkbook(data.FileName).Worksheets.First();
                        break;
                    case "Diplom":
                        data.WorksheetDiplom = new ClosedXML.Excel.XLWorkbook(data.FileName).Worksheets.First();
                        break;
                }      
            }
            catch (Exception)
            {
                MessageBox.Show("Не могу открыть XLSX файл");
                return true;
            }
            return false;
        }

        // записывает путь файла-основы
        public bool OpenOsnova()
        {
            var choofdlog = new OpenFileDialog
            {
                Title = "Выбирете файл-основу",
                Filter = "Word документ|*.docx",
                FilterIndex = 1,
                Multiselect = false
            };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                data.Osnova_flnm = choofdlog.FileName; // путь к Word файлу
                return false;
            }
            else
            {
                MessageBox.Show("Файл-основа не выбран");
                return true;
            }
        }

        // выбирает куда сохранить ворд файл и создает копию файла 
        public void OpenWord()
        {
            data.Document = null;
            data.Application = null;

            if(data.Information.Count < 1)
            {
                MessageBox.Show("Выберите файл со студентами");
                return;
            }

            if (data.Path == "") 
            {
                TakePath();
            }

            try
            {
                File.WriteAllBytes(path: data.Path, bytes: Properties.Resources.Основа2);
            }
            catch(Exception)
            {
                MessageBox.Show("Не смогло создат файл по выбраному пути. \nПопробуйте выбрать другую папку, еслди это не поможет, то зовите Витю");
                return;
            }

            try
            {
                data.Application = new Microsoft.Office.Interop.Word.Application();
                data.Document = data.Application.Documents.Open(data.Path);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                if(data.Document !=null)
                    data.Document.Close(SaveChanges: ref data.falseObj, OriginalFormat: ref data.missingObj, RouteDocument: ref data.missingObj);

                if(data.Application!=null)
                    data.Application.Quit(SaveChanges: ref data.missingObj, OriginalFormat: ref data.missingObj, RouteDocument: ref data.missingObj);

                data.Document = null;
                data.Application = null;
                File.Delete(data.Path);
                MessageBox.Show("Не могу открыть Word-файл, зовите Витю");
                return;
            }
            DropToWord();
        }
        
        // записываем из эксель файла в список с строками
        private bool FromXLSL_toForm()
        {
            if (data.Worksheet.Cell(data.Iterator, "E").Value.ToString() == "" && data.Worksheet.Cell(data.Iterator + 1, "E").Value.ToString() == "") 
            {
                MessageBox.Show("Конец файла");
                data.Iterator--;
                return false;
            }
            try
            {
                data.Information.Clear();
                /*0 - Фамилия*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "E").Value.ToString().Split(new char[] { ' ' })[0].ToString());
                /*1 - ИмяОтчество*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "E").Value.ToString().Split(new char[] { ' ' })[1].ToString() + " " + data.Worksheet.Cell(data.Iterator, "E").Value.ToString().Split(new char[] { ' ' })[2]);
                /*2 - FamilyName*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "M").Value.ToString().Split(new char[] { ' ' })[0].ToString());
                /*3 - Name*/ data.Information.Add(data.Worksheet.Cell(data.Iterator, "M").Value.ToString().Split(new char[] { ' ' })[1].ToString());

                /*4 - ДатаРождения*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "F").Value.ToString().Split(new char[] { ' ' })[0].ToString());
                
                if (data.Worksheet.Cell(data.Iterator, "S").Value.ToString() == "Магістр")
                {
                    data.proff = Find_proff(data.Worksheet_Baza_Mg, data.Worksheet.Cell(data.Iterator, "Y").Value.ToString().Split(new char[] { ' ' })[0]);
                    /*5 - Квалификация*/data.Information.Add(data.Worksheet_Baza_Mg.Cell(data.proff, "C").Value.ToString());
                    /*6 - УровеньКвалификации.Text*/data.Information.Add(data.Worksheet_Baza_Mg.Cell(data.proff, "D").Value.ToString());
                    /*7 - ДлительностьОбучения.Text*/data.Information.Add("1 рік 5 місяців, денна форма навчання (90.00 кредитів ЄКТС) @1 year 5 months, full-time form of studies (90.00 credits ECTS)");
                    /*8 - ТребованияК_Вступлению.Text*/data.Information.Add("Перший(бакалаврський) рівень вищої освіти; освітньо - кваліфікаційний рівень спеціаліст. Вступ здійснюється за результатами вступних випробовувань.@The first(bachelor's) level of higher education; educational qualification level specialist.The introduction is based on the results of entrance examinations");

                    /*9 - АкадемическиеПрава*/data.Information.Add(data.Worksheet_Baza_Mg.Cell(data.proff, "E").Value.ToString());
                    /*10 - ПроффесиональныеПрава*/data.Information.Add(data.Worksheet_Baza_Mg.Cell(data.proff, "F").Value.ToString());

                    /*11 - БазовыйДокумент*/data.Information.Add(data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[0] + "/Diploma of Bachelor " + (data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString()[0] == ' ' ? data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString().Substring(1).Replace(" ", " № ") : data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString().Replace(" ", " № ")));
                    /*12 - */data.Information.Add("Тема магістерської дипломної роботи");
                    /*13 - */data.Information.Add("The theme of master’s paper");
                }
                else
                {
                    data.proff = Find_proff(data.Worksheet_Baza_Bk, data.Worksheet.Cell(data.Iterator, "Y").Value.ToString().Split(new char[] { ' ' })[0]);
                    /*5 - Квалификация*/data.Information.Add(data.Worksheet_Baza_Bk.Cell(data.proff, "C").Value.ToString());
                    /*6 - УровеньКвалификации*/data.Information.Add(data.Worksheet_Baza_Bk.Cell(data.proff, "D").Value.ToString());
                    /*7 - ДлительностьОбучения*/data.Information.Add("Тут она по хитрому считается@уточнить!!!!");
                    /*8 - ТребованияК_Вступлению*/data.Information.Add("И тут всё узнать@!!!!!!");

                    /*9 - АкадемическиеПрава*/data.Information.Add(data.Worksheet_Baza_Bk.Cell(data.proff, "E").Value.ToString());
                    /*10 - ПроффесиональныеПрава*/data.Information.Add(data.Worksheet_Baza_Bk.Cell(data.proff, "F").Value.ToString());

                    /*11 - БазовыйДокумент*/data.Information.Add(data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[0].ToString() + "/Atestat of complete secondary education " + (data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString()[0] == ' ' ? data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString().Substring(1).Replace(" ", " № ") : data.Worksheet.Cell(row: data.Iterator, column: "AO").Value.ToString().Split(new char[] { ';' })[1].ToString().Replace(" ", " № ")));
                    /*12 - */data.Information.Add("Тема дипломної роботи");
                    /*13 - */data.Information.Add("Thema of diploma work");
                }

                /*14 - ОбластьЗнаний*/data.Information.Add(data.Worksheet_Baza_Bk.Cell(data.proff, "B").Value.ToString());

                /*15 - ФормаОбучения*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "U").Value.ToString()=="Заочна"?"Заочна@Part-time" : "Денна@Full-time");

                /*16 - ДатыОбучения*/data.Information.Add(data.Worksheet.Cell(data.Iterator, "P").Value.ToString().Split(new char[] { ' ' })[0] + "-" + data.Worksheet.Cell(data.Iterator, "Q").Value.ToString().Split(new char[] { ' ' })[0]);
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка с записью " + e.Message);
                data.Worksheet = null;
                return false;
            }
            return true;
        }

        public List<string> InformationReturn() => data.Information;
        
        // ищет номер специальности
        private int Find_proff(ClosedXML.Excel.IXLWorksheet baza, string kod)
        {
            int iterator = 2;
            string a = baza.Cell(iterator, 1).Value.ToString();

            do
            {
                if (a == kod)
                {
                    break;
                }
                else { iterator++; }
            } while (baza.Cell(data.Iterator, 1).Value.ToString() == "");
            return data.Iterator;
        }

        // записывает в ворд файл
        public void DropToWord()
        {
            ThemeOfDiploma();
            try
            {
                data.Bookmarks = data.Document.Bookmarks;
            }
            catch(Exception e)
            {
                MessageBox.Show("Не смогло создать закладки. Зовите Витю");
                //rrorLog(e);
                return;
            }
            
            try
            {
                data.Bookmarks[24].Range.Text = data.Information[15].Split(new char[] { '@' })[0] + "\n" + data.Information[15].Split(new char[] { '@' })[1];
                data.Bookmarks[23].Range.Text = data.Information[0];
                data.Bookmarks[22].Range.Text = data.Information[2];
                data.Bookmarks[21].Range.Text = data.Information[6].Split(new char[] { '@' })[0] + "\n" + data.Information[6].Split(new char[] { '@' })[1];
                data.Bookmarks[20].Range.Text = data.Information[8].Split(new char[] { '@' })[0] + "\n" + data.Information[8].Split(new char[] { '@' })[1];
                data.Bookmarks[19].Range.Text = data.Information[12];
                data.Bookmarks[18].Range.Text = data.Information[13];
                data.Bookmarks[17].Range.Text = data.Information[10].Split(new char[] { '@' })[0] + "\n" + data.Information[10].Split(new char[] { '@' })[1];
                data.Bookmarks[16].Range.Text = ""; // от какого додаток
                data.Bookmarks[15].Range.Text = ""; // от какого диплом
                data.Bookmarks[14].Range.Text = data.Information[5].Split(new char[] { '@' })[0] + "\n" + data.Information[5].Split(new char[] { '@' })[1];
                data.Bookmarks[13].Range.Text = data.Information[1];
                data.Bookmarks[12].Range.Text = data.Information[3];
                data.Bookmarks[11].Range.Text = ""; // додаток
                data.Bookmarks[10].Range.Text = data.Information[7].Split(new char[] { '@' })[0] + "\n" + data.Information[7].Split(new char[] { '@' })[1];
                data.Bookmarks[9].Range.Text = data.WorksheetDiplom.Cell(data.n, 2).Value.ToString().Split(new char[] { '@' })[0] + "/";
                data.Bookmarks[8].Range.Text = data.WorksheetDiplom.Cell(data.n, 2).Value.ToString().Split(new char[] { '@' })[1];
                data.Bookmarks[7].Range.Text = ""; //диплом
                data.Bookmarks[6].Range.Text = data.Information[16];
                data.Bookmarks[5].Range.Text = data.Information[4];
                data.Bookmarks[4].Range.Text = data.Information[14].Split(new char[] { '@' })[0] + "\n" + data.Information[14].Split(new char[] { '@' })[1];
                data.Bookmarks[3].Range.Text = "Диплом/Diploma";
                data.Bookmarks[2].Range.Text = data.Information[11];
                data.Bookmarks[1].Range.Text = data.Information[9].Split(new char[] { '@' })[0] + "\n" + data.Information[9].Split(new char[] { '@' })[1];
            }
            catch (Exception e)
            {
                data.Information.Clear();
                data.Document.Close(SaveChanges: ref data.falseObj, OriginalFormat: ref data.missingObj, RouteDocument: ref data.missingObj);
                data.Application.Quit(SaveChanges: ref data.missingObj, OriginalFormat: ref data.missingObj, RouteDocument: ref data.missingObj);
                data.Document= null;
                data.Application = null;
                MessageBox.Show(e.Message+" Не могу записать в Word-файл. \n(Или файл-основа стал неправильный и там не хватает закладок или зовите Витю)");
                return;
            }

            MakeTable();

            data.Application.Visible = true;
            data.Iterator++;
            data.Information.Clear();
            FromXLSL_toForm();
        }

        // открывает сводную ведомость и ищет там нужного студента
        public void MakeTable()
        {
            if (data.WorksheetRaiting == null) 
            {
                MessageBox.Show("Вы не выбрали файл с оценками");
                return;
            }

            int cell_row = 5;
        
            try
            {
                do
                {
                    if (data.WorksheetRaiting.Cell(cell_row, 3).Value.ToString() == data.Information[0] + " " + data.Information[1])
                    {
                        break;
                    }
                    else { cell_row++; }
                } while (data.WorksheetRaiting.Cell(cell_row, 3).Value.ToString()!=""&&data.WorksheetRaiting.Cell(cell_row+1, 3).Value.ToString()!="");
                // оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
                //table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Не нашло " + data.Information[0] + " " + data.Information[1] + "в сводной ведомости");
                return;
            }
            From_Excel_to_word(cell_row, data.Document.Tables[4]);
        }

        // достаёт оценки из сводной ведомости и записывает их в таблицу в выходном файле
        private void From_Excel_to_word(int cell_row, Table table)
        {
            int iterator = 1;
            int i = 3;
            int j = 4;

            int prev_page = table.Cell(1, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];

            // новая страница
            while (data.WorksheetRaiting.Cell(cell_row, j).Value.ToString() != "" || data.WorksheetRaiting.Cell(cell_row, j + 1).Value.ToString() != "")
            {
                table.Rows.Add(data.missingObj); //добавляем новую строку в таблицу - это надо делать всегда

                if (table.Cell(i, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber] != prev_page) // проверка на новую страницу
                {
                    for (int n = 1; n < table.Columns.Count + 1; n++)
                    {
                        table.Cell(i, n).Range.Text = table.Cell(1, n).Range.Text;
                        table.Cell(i, n).Range.Bold = 0;
                    }

                    i++;
                    prev_page++;

                    for (int n = 1; n < table.Columns.Count + 1; n++)
                    {
                        table.Cell(i, n).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        table.Cell(i, n).Range.Bold = 1;
                    }
                    continue;
                }

                // средний балл
                if (data.WorksheetRaiting.Cell(row: 2, column: j).Value.ToString() == "середній бал" || data.WorksheetRaiting.Cell(row: 2, column: j).Value.ToString() == "Середній бал")
                {
                    table.Cell(i, 1).Range.Text = "";
                    table.Cell(i, 2).Range.Text = "Підсумкова оцінка / Total mark and rank";
                    table.Cell(i, 2).Range.Bold = 1;
                    table.Cell(i, 3).Range.Text = "";
                    table.Cell(i, 4).Range.Text = "";
                    table.Cell(i, 5).Range.Text = data.WorksheetRaiting.Cell(cell_row, j).Value.ToString().Substring(0, 5);
                    table.Cell(i, 6).Range.Text = "";
                    table.Cell(i, 7).Range.Text = "";

                    StyleMethod(table);
                    break;
                }

                // выделенное 
                // сравнить работу с | и ||
                if (data.WorksheetRaiting.Cell(2, j).Value.ToString() == "Курсові роботи / Academic year papers" || data.WorksheetRaiting.Cell(2, j).Value.ToString() == "Практики / Practical training" || data.WorksheetRaiting.Cell(2, j).Value.ToString() == "Атестація / Certification")
                {
                    table.Cell(i, 1).Range.Text = "";
                    table.Cell(i, 2).Range.Text = data.WorksheetRaiting.Cell(2, j).Value.ToString(); // предмет
                    for (int n = 3; n < table.Columns.Count + 1; n++)
                    {
                        table.Cell(i, n).Range.Text = "";
                    }

                    table.Cell(i, 2).Range.Bold = 1;
                    table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // меняет выравнивание
                    j++; i++;
                    continue;
                }
                // всё остальное
                else
                {
                    table.Cell(i, 1).Range.Text = iterator.ToString() + '.'; // номер
                    table.Cell(i, 2).Range.Text = data.WorksheetRaiting.Cell(2, j).Value.ToString(); // предмет
                    table.Cell(i, 3).Range.Text = (data.WorksheetRaiting.Cell(4, j).Value.ToString() == "") ? "" : (Convert.ToDouble(data.WorksheetRaiting.Cell(4, j).Value) / 30).ToString();   // кредиты
                    table.Cell(i, 4).Range.Text = data.WorksheetRaiting.Cell(4, j).Value.ToString();   // часы
                    table.Cell(i, 5).Range.Text = data.WorksheetRaiting.Cell(cell_row, j).Value.ToString();   // баллы
                    table.Cell(i, 6).Range.Text = ConvertToLetters(data.WorksheetRaiting.Cell(cell_row, j).Value.ToString())[0];   // 
                    table.Cell(i, 7).Range.Text = ConvertToLetters(data.WorksheetRaiting.Cell(cell_row, j).Value.ToString())[1];   // 


                    table.Cell(i, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    table.Cell(i, 4).Range.Bold = 1;
                    table.Cell(i, 5).Range.Bold = 1;
                    for (int n = 3; n < table.Columns.Count + 1; n++)
                    {
                        table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                }

                StyleMethod(table);
                StyleBold(table, i);

                iterator++;
                j++; i++;
            }
        }
        private void StyleMethod(Table table)
        {
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 8;
            table.Rows.HeightRule = 0;
        }
        private void StyleBold(Table table, int i)
        {
            for (int x = 1; x < 8; x++)
                table.Cell(i, x).Range.Bold = 0;
        }

        // находит номер строки для темы диплома
        private void ThemeOfDiploma()
        {
            if (data.WorksheetDiplom == null)
            {
                MessageBox.Show("Вы не выбрали файл с темой диплома");
                return;
            }
            
            try
            {
                do
                {
                    if (data.WorksheetDiplom.Cell(data.n, 1).Value.ToString() == data.Information[0] + " " + data.Information[1])
                    {
                        break;
                    }
                    else { data.n++; }
                } while (data.WorksheetDiplom.Cell(data.n,1).Value.ToString()!=""&& data.WorksheetDiplom.Cell(data.n+1, 1).Value.ToString() != "");
                //оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
                //table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.TargetSite.Name + "; не нашло нужного человека");
                data.n = 1;
                return;
            }
        }

        // метод для таблицы с оценками; возвращает национальную шкалу и ETCS оценку
        private string[] ConvertToLetters(string str)
        {
            switch (Convert.ToInt32(str))
            {
                case int n when (n>=90):
                    return new string[] { "Відммінно/Excelent", "A" };
                    
                case int n when (n>=82):
                    return new string[] { "Добре/Good", "B" };
                    
                case int n when (n>=74):
                    return new string[] { "Добре/Good", "С" };
                    
                case int n when (n>=64):
                    return new string[] { "Задовільно/Satisfactory", "D" };
                    
                case int n when (n>=60):
                    return new string[] { "Задовільно/Satisfactory", "E" };
                    
                case int n when (n>=35):
                    return new string[] { "Незадовільно/Unsatisfactory", "FX" };
                    
                case int n when (n>=1):
                    return new string[] { "Незадовільно/Unsatisfactory", "F" };
                    
                default:
                    return new string[] { "-/-", "-" };
            }
        }

        public bool Left()
        {
            if (data.Worksheet != null)
            {
                if (data.Iterator != 2)
                {
                    data.Iterator--;
                    //data.Information.Clear();
                    FromXLSL_toForm();
                    return false;
                }
                MessageBox.Show("Это первый студент в этом файле");
                return true;
            }
            MessageBox.Show("Вы ещё не открыли файл");
            return true;
        }
        public bool Right()
        {
            if(data.Worksheet!=null)
            {
                data.Iterator++;
                //data.Information.Clear();
                if (FromXLSL_toForm())
                { return false; }

                return false;
            }
            MessageBox.Show("Вы ещё не открыли файл");
            return true;
        }

        public void FirstIterator()
        {
            data.Iterator = 2;
        }

        public void TakePath()
        {
            if (data.Information.Count > 0) 
            {
                try
                {
                    using (var fbd = new FolderBrowserDialog())
                    {
                        fbd.RootFolder = Environment.SpecialFolder.Desktop;
                        fbd.Description = "Куда сохранить файл?";

                        if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        {
                            data.Path = fbd.SelectedPath + "\\" + data.Information[0] + " " + data.Information[1] + ".doc";
                        }
                        else
                        {
                            MessageBox.Show("Вы не выбрали, куда сохранять файл");
                            return;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Не смогло выбрать, куда сохрянять файл. Зовите Витю");
                }
            }
            else
            {
                MessageBox.Show("Сначала выберите файл со студентами");
            }
        }
    }
}