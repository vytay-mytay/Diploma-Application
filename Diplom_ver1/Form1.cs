using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Diplom_ver1
{
    public partial class Form1 : Form
    {
        DataWorker dataWorker = new DataWorker();

        public Form1()
        {
            InitializeComponent();
            DoubleBuffered = true;
            //if (dataWorker.OpenOsnova())
            //    Close();
        }

        private void BTN_ОткрытьXLSX_Click(object sender, EventArgs e)
        {
            dataWorker.FirstIterator();
            if (dataWorker.OpenXLSX("Information"))
                return;

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
            TB_ОбластьЗнаний.Text = liststr[14];
            TB_ФормаОбучения.Text = liststr[15];
            TB_ДатыОбучения.Text = liststr[16];
        }
        
        private void BTN_СохранитьВорд_Click(object sender, EventArgs e)
        {
            dataWorker.OpenWord();
            if (dataWorker.InformationReturn() == null)
                ToForm(dataWorker.InformationReturn());
        }

        private void BTN_ФайлОценки_Click(object sender, EventArgs e)
        {
            dataWorker.OpenXLSX("Raiting");
        }

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
        }

        private void BTN_ПутьСохран_Click(object sender, EventArgs e)
        {
            dataWorker.TakePath();
        }
                       
    }                           
}                               