using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using ClosedXML;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Diplom_ver1
{
    public partial class Form1 : Form
    {
        //object misingObj = Missing.Value;

        //Word.Application application = null;                
        //Document document = null;                      
        //ClosedXML.Excel.XLWorkbook workbook = null;         
        //ClosedXML.Excel.IXLWorksheet worksheet = null;

        //Document osnova = null;
        //Word.Application osnova_app = null;
        //private string osnova_flnm = "";

        //object missingObj = Missing.Value;
        //object trueObj = true;
        //object falseObj = false;


        //private string Filename = "";
        //private int Iterator;

        //ClosedXML.Excel.XLWorkbook workbookRaiting = null;
        //ClosedXML.Excel.IXLWorksheet worksheetRaiting = null;
        //private string FilenameRaiting = "";

        //ClosedXML.Excel.XLWorkbook workbookDiplom = null;
        //ClosedXML.Excel.IXLWorksheet worksheetDiplom = null;
        //private string FilenameDiplom = "";
        //private int n = 1;

        //ClosedXML.Excel.IXLWorksheet worksheet_Baza_Bk = new ClosedXML.Excel.XLWorkbook("Бакалавр база.xlsx").Worksheets.First();
        //ClosedXML.Excel.IXLWorksheet worksheet_Baza_Mg = new ClosedXML.Excel.XLWorkbook("Магистр база.xlsx").Worksheets.First();


        //private int proff = 0;


        DataWorker dataWorker = new DataWorker();


        public Form1()
        {
            InitializeComponent();
            DoubleBuffered = true;
            if (dataWorker.OpenOsnova())
                Close();
        }

        private void BTN_ОткрытьXLSX_Click(object sender, EventArgs e)
        {
            dataWorker.FirstIterator();
            dataWorker.OpenXLSX("Information");
            if (dataWorker.InformationReturn() != null)
                ToForm(dataWorker.InformationReturn());
        }

        private void ToForm(List<string> liststr)
        {
            TB_Фамилия.Text = liststr[0];
            TB_ИмяОтчество.Text = liststr[1];
            TB_FamilyName.Text = liststr[2];
            TB_Name.Text = liststr[3];
            TB_ДатаРождения.Text = liststr[4];
            TB_Квалификация.Text = liststr[5];
            TB_УровеньКвалификации.Text = liststr[6];
            TB_ДлительностьОбучения.Text = liststr[7];
            TB_ТребованияК_Вступлению.Text = liststr[8];
            TB_АкадемическиеПрава.Text = liststr[9];
            TB_ПроффесиональныеПрава.Text = liststr[10];
            TB_БазовыйДокумент.Text = liststr[11];
            TB_ОбластьЗнаний.Text = liststr[12];
            TB_ФормаОбучения.Text = liststr[13];
            TB_ДатыОбучения.Text = liststr[14];
        }
        
        private void BTN_СохранитьВорд_Click(object sender, EventArgs e)
        {
            dataWorker.OpenWord();
            //dataWorker.MakeTable();
            //dataWorker.DropToWord();
            //if (workbook != null)
            //{
            //    OpenWord();
            //}
            //else
            //{
            //    MessageBox.Show("Сначала выбирите XLSX файл");
            //}
        }
        //private void OpenWord()
        //{
        //    document = null;
        //    application = null;

        //    try
        //    {
        //        string fileName = "";
        //        using (var fbd = new FolderBrowserDialog())
        //        {
        //            fbd.RootFolder = Environment.SpecialFolder.Desktop;
        //            fbd.Description = "Куда сохранить файл?";

        //            if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
        //            {
        //                fileName = fbd.SelectedPath + "\\" + TB_Фамилия.Text + " " + TB_ИмяОтчество.Text + ".doc";
        //            }
        //        }

        //        File.WriteAllBytes(path: fileName, bytes: File.ReadAllBytes(osnova_flnm));

        //        application = new Word.Application();
        //        document = application.Documents.Open(fileName);

        //        DropToWord();
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message);
        //        if(document !=null)
        //            document.Close(SaveChanges: ref falseObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);

        //        if(application!=null)
        //            application.Quit(SaveChanges: ref missingObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);

        //        document = null;
        //        application = null;
        //        MessageBox.Show("Не могу открыть Word-файл");
        //    }
        //}

        //private void DropToWord()
        //{
        //    try
        //    {
        //        ThemeOfDiploma();

        //        var wBookmarks = document.Bookmarks;
                

        //        wBookmarks[22].Range.Text = TB_ФормаОбучения.Text + (worksheet.Cell(Iterator, "U").Value.ToString() == "Денна" ? "/Full-time" : "/Part-time").ToString();
        //        wBookmarks[21].Range.Text = TB_Фамилия.Text;
        //        wBookmarks[20].Range.Text = TB_FamilyName.Text;
        //        wBookmarks[19].Range.Text = TB_УровеньКвалификации.Text.Split(new char[] { '/' })[0] + "\n" + TB_УровеньКвалификации.Text.Split(new char[] { '/' })[1];
        //        wBookmarks[18].Range.Text = TB_ТребованияК_Вступлению.Text.Split(new char[] { '/' })[0] + "\n" + TB_ТребованияК_Вступлению.Text.Split(new char[] { '/' })[1];
        //        wBookmarks[17].Range.Text = TB_ПроффесиональныеПрава.Text.Split(new char[] { '/' })[0] + "\n" + TB_ПроффесиональныеПрава.Text.Split(new char[] { '/' })[1];
        //        wBookmarks[16].Range.Text = ""; // от какого додаток
        //        wBookmarks[15].Range.Text = ""; // от какого диплом
        //        wBookmarks[14].Range.Text = TB_Квалификация.Text.Split(new char[] { '/' })[0] + "\n" + TB_Квалификация.Text.Split(new char[] { '/' })[1]; //квалификация
        //        wBookmarks[13].Range.Text = TB_ИмяОтчество.Text;
        //        wBookmarks[12].Range.Text = TB_Name.Text;
        //        wBookmarks[11].Range.Text = ""; // додаток
        //        wBookmarks[10].Range.Text = TB_ДлительностьОбучения.Text.Split(new char[] { '/' })[0] + "\n" + TB_ДлительностьОбучения.Text.Split(new char[] { '/' })[1];
        //        wBookmarks[9].Range.Text = worksheetDiplom.Cell(n, 2).Value.ToString().Split(new char[] { '/' })[0].ToString() + "/";
        //        wBookmarks[8].Range.Text = worksheetDiplom.Cell(n, 2).Value.ToString().Split(new char[] { '/' })[1].ToString();
        //        wBookmarks[7].Range.Text = ""; //диплом
        //        wBookmarks[6].Range.Text = TB_ДатыОбучения.Text; 
        //        wBookmarks[5].Range.Text = TB_ДатаРождения.Text; 
        //        wBookmarks[4].Range.Text = TB_ОбластьЗнаний.Text.Split(new char[] { '/' })[0] + "\n" + TB_ОбластьЗнаний.Text.Split(new char[] { '/' })[1];
        //        wBookmarks[3].Range.Text = TB_ТипДиплома.Text; 
        //        wBookmarks[2].Range.Text = TB_БазовыйДокумент.Text;
        //        wBookmarks[1].Range.Text = TB_АкадемическиеПрава.Text.Split(new char[] { '/' })[0] + "\n" + TB_АкадемическиеПрава.Text.Split(new char[] { '/' })[1];

        //        MakeTable();

        //        application.Visible = true;
        //        Iterator++;
        //        FromXLSL_toForm( worksheet);
        //    }
        //    catch (Exception e)
        //    {
        //        document.Close(SaveChanges: ref falseObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);
        //        application.Quit(SaveChanges: ref missingObj, OriginalFormat: ref missingObj, RouteDocument: ref missingObj);
        //        document = null;
        //        application = null;
        //        MessageBox.Show("Не могу записать в Word-файл " + e.Message);
        //        FromXLSL_toForm( worksheet);
        //        return;
        //    }
        //}

        private void BTN_ФайлОценки_Click(object sender, EventArgs e)
        {
            dataWorker.OpenXLSX("Raiting");
            

            //var choofdlog = new OpenFileDialog
            //{
            //    Filter = "Excel Лист|*.xlsx",
            //    FilterIndex = 1,
            //    Multiselect = true
            //};

            //if (choofdlog.ShowDialog() == DialogResult.OK)
            //{
            //    FilenameRaiting = choofdlog.FileName; // путь к Excel файлу
            //}

            //if (Check(FilenameRaiting))
            //{
            //    return;
            //}

            //workbookRaiting = null;
            //worksheetRaiting = null;

            //try
            //{
            //    workbookRaiting = new ClosedXML.Excel.XLWorkbook(FilenameRaiting);
            //    worksheetRaiting = workbookRaiting.Worksheets.First();
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Не могу открыть файл");
            //    return;
            //}
        }

        //private void MakeTable()
        //{
        //    if (workbookRaiting == null || worksheetRaiting == null) 
        //    {
        //        MessageBox.Show("Выберите файл с оценками");
        //        return;
        //    }

        //    int i = 5;
        //    Word.Table table = document.Tables[4];  // берём таблицу из документа под таким номером

        //    try
        //    {
        //        bool exit = true;
        //        do
        //        {
        //            string a = worksheetRaiting.Cell(i, 3).Value.ToString();
        //            string b = TB_Фамилия.Text + " " + TB_ИмяОтчество.Text;
        //            if (a == b)
        //            {
        //                exit = false;
        //            }
        //            else { i++; }
        //        } while (exit);
        //        // х*й его знает как,но оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
        //        //table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.TargetSite.Name + "; не нашло нужного человека");                
        //        return;
        //    }
        //    From_Excel_to_word(i, table);
        //}

        //private void From_Excel_to_word (int cell_row, Table table)
        //{
        //    int iterator = 1;
        //    int i = 3;
        //    int j = 4;

        //    int prev_page = table.Cell(1, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];

        //    // новая страница
        //    while (worksheetRaiting.Cell(cell_row, j).Value.ToString() != "" || worksheetRaiting.Cell(cell_row, j + 1).Value.ToString() != "") 
        //    {
        //        table.Rows.Add(misingObj);

        //        if (table.Cell(i, 1).Range.Information[WdInformation.wdActiveEndAdjustedPageNumber] != prev_page) // новая страница
        //        {
        //            for (int n = 1; n < table.Columns.Count + 1; n++)
        //            {
        //                table.Cell(i, n).Range.Text = table.Cell(1, n).Range.Text;
        //                table.Cell(i, n).Range.Bold = 0;
        //            }

        //            i++;
        //            prev_page++;

        //            for (int n = 1; n < table.Columns.Count + 1; n++) 
        //            {
        //                table.Cell(i, n).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        //                table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //                table.Cell(i, n).Range.Bold = 1;
        //            }

        //            continue;
        //        }
                

        //        // средний балл
        //        if (worksheetRaiting.Cell(row: 2, column: j).Value.ToString() == "середній бал" || worksheetRaiting.Cell(row: 2, column: j).Value.ToString() == "Середній бал")
        //        {
        //            table.Cell(i, 1).Range.Text = "";
        //            table.Cell(i, 2).Range.Text = "Підсумкова оцінка / Total mark and rank";
        //            table.Cell(i, 2).Range.Bold = 1;
        //            table.Cell(i, 3).Range.Text = "";
        //            table.Cell(i, 4).Range.Text = "";
        //            table.Cell(i, 5).Range.Text = worksheetRaiting.Cell(cell_row, j).Value.ToString().Substring(0,5);
        //            table.Cell(i, 6).Range.Text = "";
        //            table.Cell(i, 7).Range.Text = "";

        //            StyleMethod(table);

        //            break;
        //        }

        //        // выделенное 
        //        if (worksheetRaiting.Cell(2, j).Value.ToString() == "Курсові роботи / Academic year papers" | worksheetRaiting.Cell(2, j).Value.ToString() == "Практики / Practical training" | worksheetRaiting.Cell(2, j).Value.ToString() == "Атестація / Certification")
        //        {
        //            table.Cell(i, 1).Range.Text = "";
        //            table.Cell(i, 2).Range.Text = worksheetRaiting.Cell(2, j).Value.ToString(); // предмет
        //            for (int n = 3; n < table.Columns.Count + 1; n++) 
        //            {
        //                table.Cell(i, n).Range.Text = "";
        //            }
                    
        //            table.Cell(i, 2).Range.Bold = 1;
        //            table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; // меняет выравнивание
        //            j++;
        //            i++;
        //            continue;
        //        }
        //        // всё остальное
        //        else
        //        {
        //            table.Cell(i, 1).Range.Text = iterator.ToString() + '.'; // номер
        //            table.Cell(i, 2).Range.Text = worksheetRaiting.Cell(2, j).Value.ToString(); // предмет
        //            table.Cell(i, 3).Range.Text = (worksheetRaiting.Cell(4, j).Value.ToString() == "") ? "" : (Convert.ToDouble(worksheetRaiting.Cell(4, j).Value) / 30).ToString();   // кредиты
        //            table.Cell(i, 4).Range.Text = worksheetRaiting.Cell(4, j).Value.ToString();   // часы
        //            table.Cell(i, 5).Range.Text = worksheetRaiting.Cell(cell_row, j).Value.ToString();   // баллы
        //            table.Cell(i, 6).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, j).Value.ToString())[0];   // 
        //            table.Cell(i, 7).Range.Text = ConvertToLetters(worksheetRaiting.Cell(cell_row, j).Value.ToString())[1];   // 


        //            table.Cell(i, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        //            table.Cell(i, 4).Range.Bold = 1;
        //            table.Cell(i, 5).Range.Bold = 1;
        //            for (int n = 3; n < table.Columns.Count + 1; n++)
        //            {
        //                table.Cell(i, n).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            }
        //        }

        //        StyleMethod(table);
        //        StyleBold(table, i);

        //        iterator++;
        //        j++;
        //        i++;
        //    }
        //}

        //private void StyleMethod(Table table)
        //{
        //    table.Range.Font.Name = "Times New Roman";
        //    table.Range.Font.Size = 8;
        //    table.Rows.HeightRule = 0;
        //}
        //private void StyleBold(Table table,int i)
        //{
        //    for (int x = 0; x < 7; x++)
        //        table.Cell(i, x).Range.Bold = 0;
        //}

        //private string[] ConvertToLetters(string str)
        //{
        //    if (Convert.ToInt32(str)>=90)
        //    {
        //        return new string[] { "Відммінно", "A" };
        //    }
        //    else if(Convert.ToInt32(str) >= 82)
        //    {
        //        return new string[] { "Добре", "B" };
        //    }
        //    else if (Convert.ToInt32(str) >= 74)
        //    {
        //        return new string[] { "Добре", "С" };
        //    }
        //    else if (Convert.ToInt32(str) >= 64)
        //    {
        //        return new string[] { "Задовільно", "D" };
        //    }
        //    else if (Convert.ToInt32(str) >= 60)
        //    {
        //        return new string[] { "Задовільно", "E" };
        //    }
        //    else if (Convert.ToInt32(str) >= 35)
        //    {
        //        return new string[] { "Незадовільно", "FX" };
        //    }
        //    else if (Convert.ToInt32(str) >= 35)
        //    {
        //        return new string[] { "Незадовільно", "FX" };
        //    }
        //    else if (Convert.ToInt32(str) >= 1)
        //    {
        //        return new string[] { "Незадовільно", "F" };
        //    }
        //    return new string[] { "-", "-" };
        //}

        private void BTN_Left_Click(object sender, EventArgs e)
        {
            if (dataWorker.Left())
                return;
            if (dataWorker.InformationReturn() != null)
                ToForm(dataWorker.InformationReturn());
        }
        private void BTN_Right_Click(object sender, EventArgs e)
        {
            if (dataWorker.Right())
                return;
            if (dataWorker.InformationReturn() != null)
                ToForm(dataWorker.InformationReturn());
        }

        private void BTN_ТемаДиплома_Click(object sender, EventArgs e)
        {
            dataWorker.OpenXLSX("Diplom");

            //var choofdlog = new OpenFileDialog
            //{
            //    Filter = "Excel Лист|*.xlsx",
            //    FilterIndex = 1,
            //    Multiselect = true
            //};

            //if (choofdlog.ShowDialog() == DialogResult.OK)
            //{
            //    FilenameDiplom = choofdlog.FileName; // путь к Excel файлу
            //}

            //if (Check(FilenameDiplom))
            //{
            //    return;
            //}

            //workbookDiplom = null;
            //worksheetDiplom = null;

            //try
            //{
            //    workbookDiplom = new ClosedXML.Excel.XLWorkbook(FilenameDiplom);
            //    worksheetDiplom = workbookDiplom.Worksheets.First();
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Не могу открыть файл");
            //    return;
            //}
        }

        //private void ThemeOfDiploma()
        //{
        //    if (workbookDiplom == null || worksheetDiplom == null)
        //    {
        //        MessageBox.Show("Выберите файл с оценками");
        //        return;
        //    }
            
        //    try
        //    {
        //        bool exit = true;
        //        do
        //        {
        //            string a = worksheetDiplom.Cell(n, 1).Value.ToString();
        //            string b = TB_Фамилия.Text + " " + TB_ИмяОтчество.Text;
        //            if (a == b)
        //            {
        //                exit = false;
        //            }
        //            else { n++; }
        //        } while (exit);
        //        // х*й его знает как,но оно берёт номер страницы на котором таблица - это надо для того, чтобы сделать таблицу на следующей странице
        //        //table.Cell(1, 1).Range.Text = table.Cell(1, 1).Range.Information[Word.WdInformation.wdActiveEndAdjustedPageNumber].ToString();
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.TargetSite.Name + "; не нашло нужного человека");
        //        n = 1;
        //        return;
        //    }
        //}
    }
}