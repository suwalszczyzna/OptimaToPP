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

        private void saveXLSforPP(object sender, EventArgs e)
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

        public void ConvertXLStoCsv(object sender, EventArgs e)
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
            Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excel.ActiveSheet;
      
            try
            {
                workSheet.Cells[1, "A"] = PPColumnNames.SENDING_NUMBER;
                workSheet.Cells[1, "B"] = PPColumnNames.CUSTOMER_NAME_1;
                workSheet.Cells[1, "C"] = PPColumnNames.CUSTOMER_NAME_2;
                workSheet.Cells[1, "D"] = PPColumnNames.CUSTOMER_STREET;
                workSheet.Cells[1, "E"] = PPColumnNames.CUSTOMER_HOME_NO_1;
                workSheet.Cells[1, "F"] = PPColumnNames.CUSTOMER_HOME_NO_2;
                workSheet.Cells[1, "G"] = PPColumnNames.CUSTOMER_POSTCODE;
                workSheet.Cells[1, "H"] = PPColumnNames.CUSTOMER_CITY;
                workSheet.Cells[1, "I"] = PPColumnNames.CUSTOMER_COUNTRY;
                workSheet.Cells[1, "J"] = PPColumnNames.CUSTOMER_EMAIL;
                workSheet.Cells[1, "K"] = PPColumnNames.CUSTOMER_MOBILE_PHONE;
                workSheet.Cells[1, "L"] = PPColumnNames.CUSTOMER_PHONE;
                workSheet.Cells[1, "M"] = PPColumnNames.WEIGHT;
                workSheet.Cells[1, "N"] = PPColumnNames.CASH_ON_DELIVERY;
                workSheet.Cells[1, "O"] = PPColumnNames.NBR;
                workSheet.Cells[1, "P"] = PPColumnNames.TRANSFER_TITLE;
                workSheet.Cells[1, "R"] = PPColumnNames.COMMENTS;
                workSheet.Cells[1, "S"] = PPColumnNames.COMMENTS_2;
                workSheet.Cells[1, "T"] = PPColumnNames.PAYMENT_COMPANY;

                int row = 2; 
                foreach (Pack pack in packs)
                {
                    string street, city;
                    

                    if (!pack.PostCity.Equals(string.Format(pack.City)) && pack.PostCity.Length > 3 )
                    {
                        city = pack.PostCity;
                        street = string.Format("{0}, ul. {1}", pack.City, pack.Street);
                    }
                    else
                    {
                        street = pack.Street;
                        city = pack.City;
                    }

                    if (pack.Name.Length > 60)
                    {
                        workSheet.Cells[row, "B"] = string.Format(pack.Name.Substring(0, 60));
                        workSheet.Cells[row, "C"] = string.Format(pack.Name.Substring(61));
                    }
                    else
                    {
                        workSheet.Cells[row, "B"] = string.Format(pack.Name);
                    }
                   
                    workSheet.Cells[row, "D"] = string.Format(street);
                    workSheet.Cells[row, "E"] = string.Format(pack.NumberHome1);
                    workSheet.Cells[row, "F"] = string.Format(pack.NumberHome2);
                    workSheet.Cells[row, "G"] = string.Format(pack.ZipCode);
                    workSheet.Cells[row, "H"] = string.Format(city);
                    workSheet.Cells[row, "I"] = "Polska";
                    workSheet.Cells[row, "J"] = string.Format(pack.Email);

                    
                    pack.Phone = PhoneChecker.CleanPhoneNumber(pack.Phone);

                    // Is mobile phone number?
                    if (PhoneChecker.IsMobile(pack.Phone))
                    {
                        workSheet.Cells[row, "K"] = string.Format(pack.Phone);
                    }
                    else
                    {
                        workSheet.Cells[row, "L"] = string.Format(pack.Phone);
                    }

                     workSheet.Cells[row, "M"] = "30";

                    if (pack.Payment == "Pobranie")
                    {
                        workSheet.Range["N2", "N" + row].NumberFormat = "####.00";
                        workSheet.Cells[row, "N"] = Convert.ToDouble(pack.Total);
                        workSheet.Cells[row, "P"] = string.Format("UZNANIE Poczta Polska, {0}", pack.DocNumber);
                    }
                    workSheet.Cells[row, "R"] = string.Format("{0}", pack.DocNumber);
                    workSheet.Cells[row, "S"] = string.Format("{0}", pack.DocNumber);
                    workSheet.Cells[row, "T"] = "N";

                    row++;

                }

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
