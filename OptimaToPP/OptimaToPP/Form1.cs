using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OptimaToPP
{
    public partial class Form1 : Form
    {
        string TempPath = Path.GetTempPath();
        string XlsOpenPath, CsvSavePath, XMLsavePath;

        public Form1()
        {
            InitializeComponent();
            groupBox2.Enabled = false;
            CsvSavePath = "C:\\Users\\dsuwa\\Desktop\\tempFV.txt";
            XLSpath.Text = "";

        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFile = new OpenFileDialog();
            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                XlsOpenPath = OpenFile.FileName;
                XLSpath.Text = XlsOpenPath;
            }
        }



        public void btnConvert_Click(object sender, EventArgs e)
        {
            if (XLSpath.Text != "")
            {
                ConverterXLStoCSV.ConvertExcelToCsv(XlsOpenPath, CsvSavePath, 1);
                groupBox2.Enabled = true;
                groupBox1.Enabled = false;
            }
            else
            {
                MessageBox.Show("Ścieżka do pliku XLS nie może być pusta!");
            }
        }





    }
}
