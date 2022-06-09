using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace PartList
{

    class SpreadsheetCreator
    {
        private static SaveFileDialog sfd = new SaveFileDialog();

        private static int LinhaAtual = 2;

        static private string CleanTxt(string Text) => Regex.Replace(Text, "[^0-9a-zA-Z]+", "");
        static private IXLRange Select(IXLWorksheet ws,string Init, string Fim) => ws.Range(ws.Cell(Init).Address, ws.Cell(Fim).Address);

        static public void CreateTable(IXLWorkbook wb, IXLWorksheet ws)
        {

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

            ws.Column("A").Width = 15;
            ws.Column("B").Width = 25;
            ws.Column("C").Width = 15;
            ws.Column("D").Width = 25;
            ws.Column("E").Width = 45;
            ws.Column("F").Width = 15;
            ws.Column("G").Width = 15;
            ws.Column("H").Width = 15;
            ws.Column("I").Width = 15;
            ws.Column("J").Width = 15;

            wb.SaveAs(@"C:\temp\testeExcel.xlsx");
        }

        static public void AddLine(IXLWorksheet ws,IXLWorkbook wb, string Nome, string Serial, string Marca, string Modelo, string Obs, string Processador, string Memoria, string Hd, string Lacre, string Windows)
        {
            ws.Cell(LinhaAtual, 1).Value = CleanTxt(Nome);
            ws.Cell(LinhaAtual, 2).Value = CleanTxt(Serial);
            ws.Cell(LinhaAtual, 3).Value = CleanTxt(Marca);
            ws.Cell(LinhaAtual, 4).Value = CleanTxt(Modelo);
            ws.Cell(LinhaAtual, 5).Value = CleanTxt(Obs);
            ws.Cell(LinhaAtual, 6).Value = CleanTxt(Processador);
            ws.Cell(LinhaAtual, 7).Value = CleanTxt(Memoria);
            ws.Cell(LinhaAtual, 8).Value = CleanTxt(Hd);
            ws.Cell(LinhaAtual, 9).Value = CleanTxt(Lacre); ;
            ws.Cell(LinhaAtual, 10).Value = CleanTxt(Windows);
           
            LinhaAtual+=1;
        }
        static public void Save(IXLWorksheet ws, IXLWorkbook wb)
        {
            IXLAddress firstCell = ws.FirstCellUsed().Address;
            IXLAddress lastCell = ws.LastCellUsed().Address;
            IXLRange Range = ws.Range(firstCell, lastCell);

            Range.Clear(XLClearOptions.AllFormats);

            Range.CreateTable();
            
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                wb.SaveAs(sfd.FileName + ".xlsx");
            }
            else
            {
                MessageBox.Show("Aconteceu um erro e este arquivo não foi salvo");
            }
        }
    }
}