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

        static private IXLRange Search(IXLWorksheet ws,string Init, string Fim)
        {
            return ws.Range(ws.Cell(Init).Address, ws.Cell(Fim).Address);
        }

        static public void CreateTable()
        {
            var wb = new XLWorkbook();

            var ws = wb.Worksheets.Add("Plan1");

            ws.Cell("A1").Value = "Nome";
            ws.Cell("B1").Value = "Serial";
            ws.Cell("C1").Value = "Marca";
            ws.Cell("D1").Value = "Modelo";
            ws.Cell("E1").Value = "OBS";

            ws.Cell("F1").Value = "PROCESSADOR";
            ws.Cell("G1").Value = "MEMORIA";
            ws.Cell("H1").Value = "HD";
            ws.Cell("I1").Value = "LACRE";

            ws.Cell("J1").Value = "WINDOWS";

            Search(ws, "A1", "J1").Style.Fill.BackgroundColor = XLColor.RadicalRed;
            
            wb.SaveAs(@"C:\temp\testeExcel.xlsx");
            
            
        }
    }
}
