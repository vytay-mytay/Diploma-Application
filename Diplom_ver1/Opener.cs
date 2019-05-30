using System;
using System.Linq;
using System.Windows.Forms;

namespace Diplom_ver1
{
    class Opener
    {
        Data_Storage data = new Data_Storage();

        public void OpenXLSX()
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
                return;
            }

            try
            {
                data.Worksheet = new ClosedXML.Excel.XLWorkbook(data.FileName).Worksheets.First();
            }
            catch (Exception)
            {
                MessageBox.Show("Не могу открыть XLSX файл");
                return;
            }
        }
    }
}