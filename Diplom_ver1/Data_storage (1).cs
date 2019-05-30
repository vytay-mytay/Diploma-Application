using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace Diplom_ver1
{
    public class Data_Storage
    {
        public Application Application { get; set; }
        
        public Document Document { get; set; }

        public Bookmarks Bookmarks { get; set; }

        public ClosedXML.Excel.XLWorkbook Workbook { get; set; }
    
        public ClosedXML.Excel.IXLWorksheet Worksheet { get; set; }
    
        public ClosedXML.Excel.IXLWorksheet WorksheetRaiting { get; set; }
    
        public ClosedXML.Excel.IXLWorksheet WorksheetDiplom { get; set; }
        
        public string Osnova_flnm { get; set; } = "";
        public string FileName { get; set; } = "";


        public object missingObj = Missing.Value;
        public object trueObj = true;
        public object falseObj = false;


        private string FilenameXLSX { get; set; } = "";
        

        public int Iterator;

        public int n = 1;

        public ClosedXML.Excel.IXLWorksheet Worksheet_Baza_Bk { get; set; } = new ClosedXML.Excel.XLWorkbook("Бакалавр база.xlsx").Worksheets.First();
                                                                            
        public ClosedXML.Excel.IXLWorksheet Worksheet_Baza_Mg { get; set; } = new ClosedXML.Excel.XLWorkbook("Магистр база.xlsx").Worksheets.First();

        public int proff;

        private List<string> information = new List<string>();
        public List<string> Information { get => information; set => information = value; }
    }
}