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

        private void button1_Click(object sender, EventArgs e)
        {

            GenerateObjectFromCSV(CsvSavePath);
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

            saveFileDialog1.ShowDialog();
            SaveXlsPath = saveFileDialog1.FileName;

            GenerateObjectFromCSV(CsvSavePath);
            ExportToExcel(packs, SaveXlsPath);


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

        public void GenerateObjectFromCSV(string PathToCsv)
        {
            try
            {
                using (var streamReader = File.OpenText(PathToCsv))
                {
                    var reader = new CsvReader(streamReader);
                    reader.Configuration.Delimiter = ";";
                    reader.Configuration.RegisterClassMap<PackMap>();
                    packs = reader.GetRecords<Pack>().ToList();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("{0}", e));

            }


        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:damian.suwala@gmail.com?subject=Kontakt_Optima_PocztaPolska_Konwerter";
            proc.Start();
        }

        public void ExportToExcel(List<Pack> packs, string fileName)
        {
            // Load Excel application
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            // Create empty workbook
            excel.Workbooks.Add();

            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;



            try
            {

                workSheet.Cells[1, "A"] = "NumerNadania";
                workSheet.Cells[1, "B"] = "AdresatNazwa";
                workSheet.Cells[1, "C"] = "AdresatNazwaCd";
                workSheet.Cells[1, "D"] = "AdresatUlica";
                workSheet.Cells[1, "E"] = "AdresatNumerDomu";
                workSheet.Cells[1, "F"] = "AdresatNumerLokalu";
                workSheet.Cells[1, "G"] = "AdresatKodPocztowy";
                workSheet.Cells[1, "H"] = "AdresatMiejscowosc";
                workSheet.Cells[1, "I"] = "AdresatKraj";
                workSheet.Cells[1, "J"] = "AdresatEmail";
                workSheet.Cells[1, "K"] = "AdresatMobile";
                workSheet.Cells[1, "L"] = "AdresatTelefon";
                workSheet.Cells[1, "M"] = "Masa";
                workSheet.Cells[1, "N"] = "KwotaPobrania";
                workSheet.Cells[1, "O"] = "NRB";
                workSheet.Cells[1, "P"] = "TytulPobrania";
                workSheet.Cells[1, "R"] = "Uwagi";
                workSheet.Cells[1, "S"] = "Zawartosc";
                workSheet.Cells[1, "T"] = "UiszczaOplate";

                int row = 2; // start row (in row 1 are header cells)
                foreach (Pack pack in packs)
                {
                    workSheet.Cells[row, "B"] = pack.RecipientName;
                    workSheet.Cells[row, "D"] = pack.RecipientAdress;
                    workSheet.Cells[row, "E"] = pack.RecipientNoHome;
                    workSheet.Cells[row, "F"] = pack.RecipientNoHome2;
                    workSheet.Cells[row, "G"] = pack.RecipientZIP;
                    workSheet.Cells[row, "H"] = pack.RecipientCity;
                    workSheet.Cells[row, "I"] = "Polska";
                    workSheet.Cells[row, "M"] = "30";

                    if (pack.RecipientPayment == "Pobranie")
                    {
                        workSheet.Cells[row, "N"] = pack.Total;
                        //workSheet.Cells[row, "O"] = "62150015201215200779280000";
                        workSheet.Cells[row, "P"] = string.Format("UZNANIE Poczta Polska, {0}", pack.DocNumber);
                    }
                    workSheet.Cells[row, "R"] = string.Format("{0}", pack.DocNumber);
                    workSheet.Cells[row, "S"] = string.Format("{0}", pack.DocNumber);
                    workSheet.Cells[row, "T"] = "N";

                    row++;

                }

                //fix for last bad row
                workSheet.Cells[row-1, "I"] = "";
                workSheet.Cells[row-1, "M"] = "";
                workSheet.Cells[row-1, "T"] = "";

                workSheet.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8);

                MessageBox.Show(string.Format("Zapisano \n{0}", fileName));
                this.Close();
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception",
                    "Błąd podczas zapisu pliku\n" + exception.Message,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            finally
            {
                excel.Quit();

                // Release COM objects (very important!)
                if (excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                if (workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                // Empty variables
                excel = null;
                workSheet = null;

                // Force garbage collector cleaning
                GC.Collect();
            }







        }

    }
}
