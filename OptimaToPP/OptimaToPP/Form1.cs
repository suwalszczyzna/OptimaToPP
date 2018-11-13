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

        public void convertXLStoCsv(object sender, EventArgs e)
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

                int row = 2; 
                foreach (Pack pack in packs)
                {
                    string street, city;
                    

                    if (!pack.PostCity.Equals(string.Format(pack.City)) && pack.PostCity.Length != 0 )
                    {
                        city = pack.PostCity;
                        street = string.Format("{0}, ul. {1}", pack.City, pack.Street);
                    }
                    else
                    {
                        street = pack.Street;
                        city = pack.City;
                    }
                                  
                    workSheet.Cells[row, "B"] = string.Format(pack.Name);
                    workSheet.Cells[row, "D"] = string.Format(street);
                    workSheet.Cells[row, "E"] = string.Format(pack.NumberHome1);
                    workSheet.Cells[row, "F"] = string.Format(pack.NumberHome2);
                    workSheet.Cells[row, "G"] = string.Format(pack.ZipCode);
                    workSheet.Cells[row, "H"] = string.Format(city);
                    workSheet.Cells[row, "I"] = "Polska";
                    workSheet.Cells[row, "J"] = string.Format(pack.Email);

                    // Is mobile phone number or not?
                    Boolean IsMobilePhone = MobilePhoneChecker(pack.Phone);
                    if (IsMobilePhone)
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

        private bool MobilePhoneChecker(string phone)
        {
            StringComparison comparison = StringComparison.InvariantCulture;

            string[] DirectionNumbers = new string[] {
                "50","51","450","530","531","532","533","534","535","537","538","539","570","572","574","575","576","577","578","600","601","602","603","604","605","606","607","608","609","660","661","662","663","664","665","667","668","669","691","692","693","694","695","696","697","698","721","723","724","725","726","730","731","732","733","734","735","738","781","782","784","785","788","790","791","792","794","796","797","798","880","882","885","886","887","888","889","5360","5361","5362","5363","5364","5365","5366","5367","5368","5369","5711","5712","5713","5714","5730","5731","5732","5733","5734","5739","5790","5791","5792","5793","5796","5797","5798","5799","6660","6661","6662","6663","6664","6665","6666","6667","6668","6669","6900","6901","6902","6903","6904","6905","6906","6908","6909","6991","6993","6994","6998","6999","7200","7201","7202","7203","7204","7205","7206","7207","7208","7209","7220","7221","7222","7223","7224","7225","7226","7227","7228","7229","7270","7271","7272","7273","7274","7275","7276","7277","7278","7279","7280","7281","7282","7283","7284","7285","7286","7287","7288","7289","7290","7291","7292","7293","7294","7295","7296","7298","7299","7360","7361","7362","7363","7364","7365","7366","7367","7368","7369","7370","7371","7372","7373","7374","7375","7376","7377","7378","7379","7390","7391","7392","7394","7395","7396","7397","7398","7800","7801","7807","7808","7830","7831","7832","7833","7834","7835","7836","7837","7838","7839","7861","7862","7865","7866","7867","7868","7869","7870","7871","7872","7873","7874","7875","7876","7877","7878","7879","7890","7891","7892","7893","7894","7895","7896","7897","7898","7899","7930","7931","7932","7933","7934","7935","7936","7937","7938","7939","7950","7951","7952","7953","7954","7955","7956","7957","7958","7959","7990","7991","7992","7993","7994","7995","7997","7998","7999","8810","8811","8812","8813","8814","8815","8816","8817","8818","8819","8822","8830","8831","8832","8833","8834","8835","8836","8837","8838","8839","8840","8841","8842","8843","8844","8845","8846","8847","8848","8849","57941","57942","57950","57951","57952","57953","57954","57955","57956","57957","69900","69901","69902","69903","69904","69905","69906","69907","69908","69909","69920","69921","69922","69923","69924","69925","69926","69927","69928","69929","69950","69951","69952","69953","69954","69955","69956","69957","69958","69959","69960","69961","69962","69963","69964","69965","69966","69967","69968","69969","69970","69971","69972","69973","69974","69975","69976","69977","69978","69979","72970","72971","72972","72973","72974","72975","72976","72977","72978","72979","73930","73931","73932","73933","73934","73935","73936","73937","73938","73939","73991","73992","73993","78020","78021","78022","78023","78024","78025","78026","78027","78028","78029","78608"
            };

            for(int i=0; i < DirectionNumbers.Length; i++)
            {
                if ((phone.StartsWith(DirectionNumbers[i], comparison))){

                    return false;

                } 
               
            }
            return true;
        }
    }
}
