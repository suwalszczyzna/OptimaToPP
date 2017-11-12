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

namespace OptimaToPP
{
    public partial class Form1 : Form
    {
        string TempPath = Path.GetTempPath();
        string XlsOpenPath, CsvSavePath, XMLsavePath;
        string file;
        public Form1()
        {
            InitializeComponent();
            groupBox2.Enabled = false;
            CsvSavePath = "C:\\Users\\dsuwa\\Desktop\\tempFV.csv";
            XLSpath.Text = "";

        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFile = new OpenFileDialog();
            OpenFile.Filter = "(*.xml)|*.xml";
            if (OpenFile.ShowDialog() == DialogResult.OK)
            {
                XlsOpenPath = OpenFile.FileName;
                XLSpath.Text = XlsOpenPath;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            GenerateObjectFromCSV(CsvSavePath);
            string SaveXmlPath;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                InitialDirectory = (@"C:\"),
                Title = "Zapisz plik XML",
                FileName = "wysylki.xml",
                CheckFileExists = false,
                CheckPathExists = true,
                DefaultExt = ".xml",
                Filter = "(*.xml)|*.xml",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            saveFileDialog1.ShowDialog();
            SaveXmlPath = saveFileDialog1.FileName;


          
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

        public void GenerateObjectFromCSV (string PathToCsv)
        {
            var packs = new List<Pack>();
            using (var streamReader = File.OpenText(PathToCsv))
            {
                var reader = new CsvReader(streamReader);
                reader.Configuration.Delimiter = ";";
                reader.Configuration.RegisterClassMap<PackMap>();
                packs = reader.GetRecords<Pack>().ToList();
            }
            

            file = "<?xml version=\"1.0\" encoding=\"UTF - 8\"?>" +
                "<transactions>";
            foreach (var o in packs)
            {
                string adress = string.Format($"{o.RecipientAdress} {o.RecipientNoHome} / {o.RecipientNoHome2}");
                string payment;
                if (o.RecipientPayment == "pobranie")
                {
                    payment = "Przy odbiorze (za pobraniem)";
                }
                else
                {
                    payment = o.RecipientPayment;
                }

                file += "<transaction>";

                file += "<parentId/> \n" +
                        "<Id/>\n" +
                        "<Name/>\n" +
                        "<OrderId/>\n" +
                        "<RecipientName>"+ o.RecipientName+ "</RecipientName>\n" +
                        " <RecipientPhone>508635104</RecipientPhone>" +
                        "<RecipientAdress>" +adress+ "</RecipientAdress>\n" +
                        "<RecipientZip>"+o.RecipientZIP+"</RecipientZip>\n" +
                        "<RecipientCity>"+o.RecipientCity+"</RecipientCity>\n" +
                        "<RecipientCountryCode>PL</RecipientCountryCode>\n" +
                        "<RecipientCountryName>Polska</RecipientCountryName>\n" +
                        "<Total>"+o.Total+"</Total>\n" +
                        "<PaymentType>"+payment+"</PaymentType>";
                                
                file += "</transaction>";

            }

            file += "</transactions>";
        }



    }
}
