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


namespace PartList
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();

            SpreadsheetCreator.CreateTable();

            Process.Start(new ProcessStartInfo(@"C:\temp\testeExcel.xlsx"));
        }
    }
}
