using System;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Diplom_ver1
{
    public partial class Form1 : Form
    {
        object misingObj = Missing.Value;

        Word.Application application = null;                
        Document document = null;
        Table table = null;
       
        ClosedXML.Excel.IXLWorksheet worksheet = null;
        ClosedXML.Excel.IXLWorksheet worksheetInfFtud = null;
        ClosedXML.Excel.IXLWorksheet worksheetRaiting = null;
        ClosedXML.Excel.IXLWorksheet worksheetDiplom = null;

        private string osnova_flnm = "";

        object missingObj = Missing.Value;
        object trueObj = true;
        object falseObj = false;


        private string Filename = "";
        private int Iterator = 0;
        int RaitingIterator = 0;

        private string PathStringFile = "";
        
        public Form1()
        {
            InitializeComponent();

            try
            {
                BaseFile();
            }
            catch (Exception)
            {
                MessageBox.Show("Не смогло выбрать файл-основу(Word-основу)\nПопробуйте заново выбрать файл основу или перезагрузите приложение, предварительно проверив файл-основу.");
            }
        }

        private void BTN_ОткрытьXLSX_Click(object sender, EventArgs e)
        {
            Open((sender as Button).Name);
        }

        private void Open(string btn_name)
        {
            var choofdlog = new OpenFileDialog
            {
                Filter = "Excel Лист (xlsx)|*.xlsx",
                FilterIndex = 1,
                Multiselect = true
            };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                Filename = choofdlog.FileName; // путь к Excel файлу
            }
            else
            {
                return;
            }

            if (Check(Filename))
            {
                return;
            }
            try
            {
                switch (btn_name)
                {
                    case "BTN_ИнфСтуд":
                        worksheet = new ClosedXML.Excel.XLWorkbook(Filename).Worksheets.First();
                        FromXLSL_toForm(1);
                        break;
                    case "BTN_ДокДляОбраз":
                        worksheetInfFtud = new ClosedXML.Excel.XLWorkbook(Filename).Worksheets.First();
                        FromXLSL_toForm(2);
                        break;
                    case "BTN_ФайлОценки":
                        worksheetRaiting = new ClosedXML.Excel.XLWorkbook(Filename).Worksheets.First();
                        break;
                    case "BTN_ФайлДиплом":
                        worksheetDiplom = new ClosedXML.Excel.XLWorkbook(Filename).Worksheets.First();
                        FromXLSL_toForm(4);
                        break;
                    default:
                        MessageBox.Show("Вы нажали неизвесную кнопку.");
                        break;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Не могу открыть файл" + "\n" + e.Message);
                return;
            }
        }
        private bool Check(string filename)
        {
            if (filename == "")
            {
                MessageBox.Show("Вы не открыли таблицу");
                return true;
            }
            return false;
        }
        private bool Check()
        {
            if (worksheet == null && worksheetDiplom == null && worksheetInfFtud == null && worksheetRaiting == null) 
            {
                MessageBox.Show("Откройте ещё таблицу");
                return true;
            }
            return false;
        }

        public void FromXLSL_toForm(int type = 0)
        {
            if (worksheet.Cell(Iterator, "L").Value.ToString() == "" && worksheet.Cell(Iterator + 1, "L").Value.ToString() == "") 
            {
                MessageBox.Show("Конец файла");
                Iterator--;
                return;
            }
            try
            {
                if ((worksheet != null) && (type == 1 || type == 0))
                {
                    TB_СерДип.Text = worksheet.Cell(Iterator, "I").Value.ToString().Trim();
                    TB_НомДип.Text = worksheet.Cell(Iterator, "J").Value.ToString().Trim();
                    TB_ДатДип.Text = worksheet.Cell(Iterator, "AN").Value.ToString().Trim().Split(new char[] { ' ' })[0];
                    TB_Фамилия.Text = worksheet.Cell(Iterator, "L").Value.ToString().Trim();
                    TB_Имя.Text = worksheet.Cell(Iterator, "M").Value.ToString().Trim();
                    TB_Отчество.Text = worksheet.Cell(Iterator, "N").Value.ToString().Trim();
                    TB_FamilyName.Text = worksheet.Cell(Iterator, "O").Value.ToString().Trim();
                    TB_Name.Text = worksheet.Cell(Iterator, "P").Value.ToString().Trim();
                    TB_ДатаРождения.Text = worksheet.Cell(Iterator, "R").Value.ToString().Trim().Split(new char[] { ' ' })[0];
                    TB_ТипДиплома.Text = (worksheet.Cell(Iterator, "V").Value.ToString().Trim()=="З відзнакою")? "Диплом з відзнакою/Honors degree" : "Диплом/Diploma";
                    TB_ФормаОбучения.Text = worksheet.Cell(Iterator, "AA").Value.ToString().Trim();
                    TB_ДатаКонцаУчёбы.Text = worksheet.Cell(Iterator, "AO").Value.ToString().Trim().Split(new char[] { ' ' })[0];
                }
                if ((worksheetInfFtud != null) && (type == 2 || type == 0))
                {
                    TB_ДатаНачалаУчёбы.Text = worksheetInfFtud.Cell(Iterator, "P").Value.ToString().Trim().Split(new char[] { ' ' })[0];
                    TB_БазовыйДокумент.Text = worksheetInfFtud.Cell(Iterator, "AO").Value.ToString().Trim().Split(new char[] { ';' })[0];
                    TB_СерияБазДок.Text = worksheetInfFtud.Cell(Iterator, "AO").Value.ToString().Trim().Split(new char[] { ';' })[1].Trim().Split(new char[] { ' ' })[0];
                    TB_НомерБазовогоДокумента.Text = worksheetInfFtud.Cell(Iterator, "AO").Value.ToString().Trim().Split(new char[] { ';' })[1].Trim().Split(new char[] { ' ' })[1];
                }
                if ((worksheetDiplom != null) && (type == 4 || type == 0)) 
                {
                    TB_СерДод.Text = worksheetDiplom.Cell(Iterator, "F").Value.ToString().Trim();
                    TB_НомДод.Text = worksheetDiplom.Cell(Iterator, "G").Value.ToString().Trim();
                    TB_ДатДод.Text = worksheetDiplom.Cell(Iterator, "H").Value.ToString().Trim().Split(new char[] { ' ' })[0];
                    TB_ТемаДипУкр.Text = worksheetDiplom.Cell(Iterator, "D").Value.ToString().Trim();
                    TB_ТемаДипАнгл.Text = worksheetDiplom.Cell(Iterator, "E").Value.ToString().Trim();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка с выводом информации на форму");
                worksheet = null;
                return;
            }
        }
        
     
        private void BTN_СохранитьВорд_Click(object sender, EventArgs e)
        {
            if (osnova_flnm != "") 
            {
                if (!Check())
                {
                    OpenWord();
                }
                else
                {
                    MessageBox.Show("Сначала выбирите XLSX файл");
                }
            }
            else
            {
                MessageBox.Show("Сначала выбирите Word-файл-основу");
            }
        }
        private void OpenWord()
        {
            document = null;
            application = null;

            try
            {
                File.WriteAllBytes(path: PathStringFile + TB_Фамилия.Text + " " + TB_Имя.Text + " " + TB_Отчество.Text + ".doc", bytes: File.ReadAllBytes(osnova_flnm));

                application = new Word.Application();
                document = application.Documents.Open(PathStringFile + TB_Фамилия.Text + " " + TB_Имя.Text + " " + TB_Отчество.Text + ".doc");

                DropToWord();

                //Thread TestThread = new Thread(new ThreadStart(MakeTable));
                //TestThread.Start();

                //PathStringFile = "";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                if(document !=null)
                    document.Close(SaveChanges: ref falseObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);

                if(application!=null)
                    application.Quit(SaveChanges: ref missingObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);

                document = null;
                application = null;
                MessageBox.Show("Не могу открыть Word-файл");
            }
        }

        private void DropToWord()
        {
            try
            {
                var wBookmarks = document.Bookmarks;

                wBookmarks[20].Range.Text = TB_ФормаОбучения.Text + (TB_ФормаОбучения.Text == "Денна" ? "/Full-time" : "/Part-time");
                wBookmarks[19].Range.Text = TB_Фамилия.Text;
                wBookmarks[18].Range.Text = TB_FamilyName.Text;


                wBookmarks[17].Range.Text = TB_ТипДиплома.Text;
                wBookmarks[16].Range.Text = TB_ТемаДипУкр.Text;
                wBookmarks[15].Range.Text = TB_ТемаДипАнгл.Text;


                wBookmarks[14].Range.Text = TB_СерДод.Text;
                wBookmarks[13].Range.Text = TB_СерДип.Text;
                wBookmarks[12].Range.Text = TB_СерияБазДок.Text;
                wBookmarks[11].Range.Text = TB_НомДод.Text;


                wBookmarks[10].Range.Text = TB_НомДип.Text;


                wBookmarks[9].Range.Text = TB_НомерБазовогоДокумента.Text;

                wBookmarks[8].Range.Text = TB_ДатаНачалаУчёбы.Text;

                wBookmarks[7].Range.Text = TB_ДатаКонцаУчёбы.Text;
                wBookmarks[6].Range.Text = TB_Имя.Text + " " + TB_Отчество.Text;

                wBookmarks[5].Range.Text = TB_Name.Text;
                wBookmarks[4].Range.Text = TB_ДатаРождения.Text;

                wBookmarks[3].Range.Text = TB_ДатДод.Text;
                wBookmarks[2].Range.Text = TB_ДатДип.Text;
                wBookmarks[1].Range.Text = TB_БазовыйДокумент.Text + ("TB_БазовыйДокумент.Text" == "Атестат про повну загальну середню освіту" ? "Atestat of complete secondary education" : "Somethings else");

                MakeTable();

                application.Visible = true;                
                Iterator++;
                FromXLSL_toForm(0);
            }
            catch (Exception e)
            {
                document.Close(SaveChanges: ref falseObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);
                application.Quit(SaveChanges: ref missingObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);
                document = null;
                application = null;
                MessageBox.Show("Не могу записать в Word-файл " + e.Message);
                FromXLSL_toForm(0);
                return;
            }
        }

        private void BTN_ФайлОценки_Click(object sender, EventArgs e)
        {
            Open((sender as Button).Name);
        }

        private void MakeTable()
        {
            if (worksheetRaiting == null) 
            {
                MessageBox.Show("Выберите файл с оценками");
                return;
            }

            RaitingIterator = Iterator + 2;
            table = document.Tables[5];  // берём таблицу из документа под таким номером

                // х*й его знает как,но оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
                //table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();

            try
            {
                From_Excel_to_word(RaitingIterator, table);
            }
            catch
            {
                MessageBox.Show("Не смогло записать оценки");
            }

        }

        private void From_Excel_to_word (int cell_row /*номер строки студента*/, Table table)
        {
            int iterator = 1;   // итератор для ворда - первая колонка в ворд таблице
            int i = 2;          // до этого было 3 - какая строка в таблице в ворд
            int ExCol = 5;          //                 - какая колонка в таблице в эксель
            int prev_page = table.Cell(1, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];

            // новая страница

            try
            {
                while (worksheetRaiting.Cell(1, ExCol).Value.ToString() != "") 
                {
                    table.Rows.Add(misingObj);
                
                    if (table.Cell(i, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber] != prev_page /*table.Cell(i-1, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber]/*prev_page*/) // новая страница
                    {
                        for (int n = 1; n < table.Columns.Count + 1; n++)
                        {
                            table.Cell(i, n).Range.Text = table.Cell(1, n).Range.Text;
                            table.Cell(i, n).Range.Bold = 0;
                        }
                
                        for (int n = 1; n < table.Columns.Count + 1; n++)
                        {
                            table.Cell(i, n).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            table.Cell(i, n).Range.Bold = 1;
                        }
                
                        table.Rows.Add(misingObj);
                        prev_page = table.Cell(i, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];
                        i++;
                    }
                
                    table.Cell(i, 1).Range.Text = iterator.ToString()/* + '.' + table.Cell(i, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber]*/; // номер
                    table.Cell(i, 2).Range.Text = worksheetRaiting.Cell(1, ExCol).Value.ToString(); // предмет
                    table.Cell(i, 3).Range.Text = (worksheetRaiting.Cell(2, ExCol).Value.ToString() == "") ? "" : (Convert.ToDouble(worksheetRaiting.Cell(2, ExCol).Value) / 30).ToString(); // кредиты
                    table.Cell(i, 4).Range.Text = worksheetRaiting.Cell(2, ExCol).Value.ToString();   // часы

                    if (worksheetRaiting.Cell(cell_row, ExCol).Value.ToString() != "")
                    {
                        if (worksheetRaiting.Cell(cell_row, ExCol).Value.ToString().Length > 4)          // баллы
                        {
                            table.Cell(i, 5).Range.Text = worksheetRaiting.Cell(cell_row, ExCol).Value.ToString().Remove(5);
                            table.Cell(i, 6).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, ExCol).Value.ToString().Remove(5))[0];   // за нац шкалой
                            table.Cell(i, 7).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, ExCol).Value.ToString().Remove(5))[1];   // буква
                        }
                        else
                        {
                            if(worksheetRaiting.Cell(3, ExCol).Value.ToString()=="з" || worksheetRaiting.Cell(3, ExCol).Value.ToString() == "З")
                            {
                                table.Cell(i, 6).Range.Text = "Зараховано/Counted";
                            }
                            else
                            {
                                table.Cell(i, 6).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, ExCol).Value.ToString())[0];   // за нац шкалой
                            }
                            table.Cell(i, 5).Range.Text = worksheetRaiting.Cell(cell_row, ExCol).Value.ToString(); // баллы
                            table.Cell(i, 7).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, ExCol).Value.ToString())[1];   // буква
                        }
                        table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        StyleBold(table, i);
                        table.Cell(i, 3).Range.Bold = 1;
                        table.Cell(i, 4).Range.Bold = 1;
                        table.Cell(i, 5).Range.Bold = 1;
                        table.Cell(i, 7).Range.Bold = 1;
                        for (int n = 3; n < table.Columns.Count + 1; n++)
                        {
                            table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        iterator++;
                    }
                    else
                    {
                        table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                        table.Cell(i, 2).Range.Bold = 1;
                        table.Cell(i, 1).Range.Text = "";
                    }
                    //StyleMethod(table);
                    table.Rows[i].Range.Font.Name = "Times New Roman";
                    table.Rows[i].Range.Font.Size = 8;
                    table.Rows[i].HeightRule = 0;
                    ExCol++; i++;
                }

                table.Cell(table.Rows.Count, 1).Range.Text = "";
                table.Cell(table.Rows.Count, 2).Range.Text = "Всього кредитів ЄКТС/ Total credits ECTS";
                table.Cell(table.Rows.Count, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                for (int y = 2; y <= 4; y++)
                {
                    table.Cell(table.Rows.Count, y).Range.Bold = 1;
                }
                string total = table.Cell(table.Rows.Count, 5).Range.Text;
                for (int y = 5; y <= 7 ; y++)  
                {
                    table.Cell(table.Rows.Count, y).Range.Text = "";
                }

                table.Rows.Add(misingObj);
                table.Cell(table.Rows.Count, 2).Range.Text = "Підсумкова оцінка / Total mark and rank";
                table.Cell(table.Rows.Count, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                for (int y = 2; y <= 5; y++)
                {
                    table.Cell(table.Rows.Count, y).Range.Bold = 1;
                }
                table.Cell(table.Rows.Count, 5).Range.Text = total;
                return;
            }
            catch(Exception)
            {
                MessageBox.Show("Не смогло записать оценку");
            }
        }

        private void StyleMethod(Table table)
        {
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 8;
            table.Rows.HeightRule = 0;
        }
        private void StyleBold(Table table,int i)
        {
            for (int x = 0; x <= 7; x++)
                table.Cell(i, x).Range.Bold = 0;
        }

        private string[] ConvertToLetters(string str)
        {
            switch (Convert.ToDouble(str))
            {
                case double n when (n >= 90.0):
                    return new string[] { "Відммінно/Excelent", "A" };

                case double n when (n >= 82.0):
                    return new string[] { "Добре/Good", "B" };

                case double n when (n >= 74.0):
                    return new string[] { "Добре/Good", "С" };

                case double n when (n >= 64.0):
                    return new string[] { "Задовільно/Satisfactory", "D" };

                case double n when (n >= 60.0):
                    return new string[] { "Задовільно/Satisfactory", "E" };

                case double n when (n >= 35.0):
                    return new string[] { "Незадовільно/Unsatisfactory", "FX" };

                case double n when (n >= 1.0):
                    return new string[] { "Незадовільно/Unsatisfactory", "F" };
                
                default:
                    return new string[] { "-/-", "-" };
            }
        }

        private void BTN_Left_Click(object sender, EventArgs e)
        {
            if (!Check() && Iterator > 2)  
            {
                Iterator--;
                FromXLSL_toForm();
                return;
            }
            MessageBox.Show("Это первый студент в этом файле");
        }
        private void BTN_Right_Click(object sender, EventArgs e)
        {
            if (!Check()) 
            {
                Iterator++;
                FromXLSL_toForm();
            }
        }
        
        // оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
        // table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();

        private void BTN_ФайлОснова_Click(object sender, EventArgs e)
        {
            try
            {
                BaseFile();
            }
            catch
            {
                MessageBox.Show("Не смогло выбрать файл-основу");
            }
        }
        private void BaseFile()
        {
            var choofdlog = new OpenFileDialog
            {
                Title = "Выбирете файл-основу",
                Filter = "Word документ (doc)|*.doc",
                FilterIndex = 1,
                Multiselect = false
            };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                osnova_flnm = choofdlog.FileName; // путь к Word файлу
            }
            else
            {
                osnova_flnm = "";
                MessageBox.Show("Вы не выбрали файл-основу");
            }
        }

        private void BTN_ПутьСохраненияФайла_Click(object sender, EventArgs e)
        {
            try
            {
                FilePath();
            }
            catch
            {
                MessageBox.Show("Не смогло выбрать путь для сохранения файла");
            }
        }
        private void FilePath()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.RootFolder = Environment.SpecialFolder.Desktop;
                fbd.Description = "Куда сохранить файл?";

                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    PathStringFile = fbd.SelectedPath + "\\";
                }
            }
        }

        private void BTN_ФайлДиплом_Click(object sender, EventArgs e)
        {
            Open((sender as Button).Name);
        }
        //private void FindTheme()
        //{
        //    if(worksheetDiplom!=null)
        //    {
        //        n = 1;
        //        do
        //        {
        //            if (worksheetDiplom.Cell(n, 1).Value.ToString() == TB_Фамилия.Text + " " + TB_Имя.Text)
        //            {
        //                return;
        //            }
        //            else { n++; }
        //        } while (worksheetDiplom.Cell(n, 1).Value.ToString() != "" && worksheetDiplom.Cell(n + 1, 1).Value.ToString() != "");
        //    }
        //    n = -1;
        //}

        private void BTN_ИнфСтуд_Click(object sender, EventArgs e)
        {
            Iterator = 2;
            Open((sender as Button).Name);
        }
    }
}