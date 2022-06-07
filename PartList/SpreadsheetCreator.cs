using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ClosedXML.Excel;

namespace PartList
{

    class SpreadsheetCreator
    {
        
        static public void CreateTable()
        {
            var wb = new XLWorkbook();

            var ws = wb.Worksheets.Add("Plan1");

            ws.Cell("A1").Value = "TESTE";
            ws.Cell("B1").Value = "TEST";
            ws.Cell("C1").Value = "TESE";
            ws.Cell("D1").Value = "TETE";
            ws.Cell("E1").Value = "TSTE";

            wb.SaveAs(@"C:\temp\testeExcel.xlsx");
        }
    }
}
