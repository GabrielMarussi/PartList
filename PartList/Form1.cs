using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

using ClosedXML.Excel;

namespace PartList
{
    public partial class Form1 : Form
    {
        private IXLWorkbook wb;
        private IXLWorksheet ws;

        public Form1()
        {
            InitializeComponent();

            wb = new XLWorkbook();
            ws = wb.Worksheets.Add("Plan1");

            SpreadsheetCreator.CreateTable(wb,ws);
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            SpreadsheetCreator.AddLine(ws, wb, TxtNome.Text, TxtSerial.Text, TxtMarca.Text, TxtModelo.Text, TxtObs.Text, TxtProcessador.Text, TxtMemoria.Text, TxtHd.Text, TxtLacre.Text, TxtWindows.Text);

            TxtNome.Text = "";
            TxtSerial.Text = "";
            TxtMarca.Text = "";
            TxtModelo.Text = "";
            TxtObs.Text = "";
            TxtProcessador.Text = "";
            TxtMemoria.Text = "";
            TxtHd.Text = "";
            TxtLacre.Text = "";
            TxtWindows.Text = "";
        }

        private void BtnSalvar_Click(object sender, EventArgs e) => SpreadsheetCreator.Save(ws, wb);
    }
}
