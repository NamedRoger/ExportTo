using System.Collections.Generic;
using System.Data;
using System.IO;
using ClosedXML.Excel;

namespace ExportTo
{
    public class Excel
    {
      public static void FormDataTable(DataTable data,string filePath,string fileName){
            XLWorkbook _wb = new XLWorkbook();
            _wb.AddWorksheet(data);
            if(!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            _wb.SaveAs(Path.Combine(filePath,$"{fileName}.xlsx"));
        }

        public void FromToList<T>(List<T> data, string filePath,string sheetName){

        }
    }
}