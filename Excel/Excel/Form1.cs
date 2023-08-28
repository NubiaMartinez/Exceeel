using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word= Microsoft.Office.Interop.Word;
using Excel2= Microsoft.Office.Interop.Excel;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialogo = new SaveFileDialog();
            if (dialogo.ShowDialog() !=DialogResult.OK)
            {
                return;
            }
            string ruta = dialogo.FileName;
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.Documents.Add();
            string dato = textBox1.Text;

            wordApp.Selection.TypeText(dato);
            wordApp.ActiveDocument.SaveAs2(ruta); 
    }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialogo = new SaveFileDialog();
            if (dialogo.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string ruta = dialogo.FileName;
            var excelApp = new Excel2.Application();
            excelApp.Visible = true;
            Excel2.Workbook workbook = excelApp.Workbooks.Add();
            Excel2.Worksheet worksheet = workbook.ActiveSheet;
            worksheet.Cells[1, 1] = textBox1.Text;

            workbook.SaveAs(ruta);
        }
    }
}
