using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace Excel_Okuma_ve_Listeleme
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelNesnesi = new Microsoft.Office.Interop.Excel.Application();

            if (ExcelNesnesi == null)
            {
                MessageBox.Show("Problem! Dosya Açılamadı.");
                System.Windows.Forms.Application.Exit();
            }
        }

        private Microsoft.Office.Interop.Excel.Application ExcelNesnesi = null;
 
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        string[] ConvertToStringArray(System.Array values)
        {
            string[] theArray = new string[values.Length];
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }

            return theArray;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            textBox1.Clear();
            openFileDialog1.FileName = "*.xls";
            textBox1.Text = openFileDialog1.FileName.ToString();
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Workbook theWorkbook =
                ExcelNesnesi.Workbooks.Open(
                openFileDialog1.FileName,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);
                Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);
                for (int i = 1; i <= 20; i++)
                {
                    Microsoft.Office.Interop.Excel.Range range =
                    worksheet.get_Range("A" + i.ToString(), "Z" + i.ToString());
                    System.Array myvalues = (System.Array)range.Cells.Value2;
                    string[] strArray = ConvertToStringArray(myvalues);
                    listBox1.Items.Add(strArray[2]);
                    textBox3.Text = listBox1.Items.Count.ToString();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == String.Empty)
            {
                return;
            }
            int index = listBox1.FindString(textBox2.Text);
            if (index != -1)
            {
                listBox1.SetSelected(index, true);
            }
            else
            {
                MessageBox.Show("Aranılan Kişi Bulunamadı!");
            }
        }
    }
}
