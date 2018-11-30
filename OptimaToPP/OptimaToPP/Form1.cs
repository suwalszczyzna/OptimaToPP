using CsvHelper;
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
using System.Xml;
using Microsoft.Office.Interop;
using OptimaToPP.Controllers;

namespace OptimaToPP
{
    public partial class Form1 : Form
    {
        string TempPath = Path.GetTempPath();
        string XlsOpenPath, CsvSavePath;
        List<Pack> packs = new List<Pack>();
        string tempPath = System.IO.Path.GetTempPath();
        public Form1()
        {
            InitializeComponent();
            this.Text = "Optima & PP - integrator";
            this.AutoScaleMode = AutoScaleMode.Dpi;
            groupBox2.Enabled = false;
            CsvSavePath = string.Format(@"{0}optima_to_pp_temp.csv", tempPath);
            XLSpath.Text = "";

        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFile = new OpenFileDialog();
            OpenFile.Filter = "(*.xls)|*.xls";
            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                XlsOpenPath = OpenFile.FileName;
                XLSpath.Text = XlsOpenPath;
            }
        }

        private async void saveXLSforPP(object sender, EventArgs e)
        {
            
            string SaveXlsPath;
            DateTime dateTime = DateTime.UtcNow.Date;
            string todayDate = dateTime.ToString("dd-MM-yyyy");

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                InitialDirectory = (string.Format(@"{0}",Environment.SpecialFolder.Personal)),
                Title = "Zapisz plik XLS",
                FileName = "wysylki" + todayDate,
                CheckFileExists = false,
                CheckPathExists = true,
                DefaultExt = ".xls",
                Filter = "(*.xls)|*.xls",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                groupBox2.Enabled = false;
                SaveXlsPath = saveFileDialog1.FileName;
                packs = await Converter.CSVtoListOfPacks(CsvSavePath);
                await Exporter.ToExcel(packs, SaveXlsPath);
                this.Close();
            }
            
        }
        public void ConvertXLStoCsv(object sender, EventArgs e)
        {
            if (XLSpath.Text != "")
            {
                Converter.XLStoCSV(XlsOpenPath, CsvSavePath, 1);
                groupBox2.Enabled = true;
                groupBox1.Enabled = false;
            }
            else
            {
                MessageBox.Show("Ścieżka do pliku XLS nie może być pusta!");
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:damian.suwala@gmail.com?subject=Kontakt_Optima_PocztaPolska_Konwerter";
            proc.Start();
        }

    }
}
