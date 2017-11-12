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
            //OpenFile.Filter = "(*.xls)|*.xml";
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

            GenerateObjectFromCSV(CsvSavePath);
           // File.WriteAllText(SaveXmlPath, @file);

            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(file);
            xdoc.Save(SaveXmlPath);

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
            

            file = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<transactions>\n" +
                "<range>\n" +
                "<from>2017-11-10</from>\n" +
                "<to>2017-11-11</to>\n" +
                "<sections>Sprzedane</sections>\n" +
                "</range>";
            foreach (var o in packs)
            {
                string adress = string.Format($"{o.RecipientAdress} {o.RecipientNoHome}  {o.RecipientNoHome2}");
                string payment;
                if (o.RecipientPayment == "pobranie")
                {
                    payment = "Przy odbiorze (za pobraniem)";
                }
                else
                {
                    payment = "Płatność elektroniczna";
                }

                file += "<transaction>\n";

                file += "<parentId/> \n" +
                        "<Id>35285349</Id>\n" +
                        "<Name>Wąż asenizacyjny superelastyczny 45mm</Name>\n" +
                        "<OrderId>6874226388</OrderId>\n" +
                        "<CustomerLogin>xx</CustomerLogin>\n"+
                        "<CustomerEmail>damian@valvotec.pl</CustomerEmail>\n" +
                        "<CustomerName>" + o.RecipientName + "</CustomerName>\n" +
                        "<CustomerPhone>508635104</CustomerPhone>\n" +
                        "<CustomerAddress>" + adress + "</CustomerAddress>\n" +
                        "<CustomerZip>" + o.RecipientZIP + "</CustomerZip>\n" +
                        "<CustomerCity>" + o.RecipientCity + "</CustomerCity>\n" +
                        "<CustomerCountryCode>PL</CustomerCountryCode>\n" +
                        "<CustomerCountryName>Polska</CustomerCountryName>\n" +
                        "<RecipientName>"+ o.RecipientName+ "</RecipientName>\n" +
                        "<RecipientCompanyName/>\n" +
                        "<RecipientPhone>508635104</RecipientPhone>\n" +
                        "<RecipientAdress>" +adress+ "</RecipientAdress>\n" +
                        "<RecipientZip>"+o.RecipientZIP+"</RecipientZip>\n" +
                        "<RecipientCity>"+o.RecipientCity+"</RecipientCity>\n" +
                        "<RecipientCountryCode>PL</RecipientCountryCode>\n" +
                        "<RecipientCountryName>Polska</RecipientCountryName>\n" +
                        "<InvoiceName/>\n"+
                        "<InvoiceCompanyName/>\n"+
                        "<InvoiceAddress/>\n"+
                        "<InvoiceZip/>\n"+
                        "<InvoiceCity/>\n"+
                        "<InvoiceCountryCode/>\n"+
                        "<InvoiceCountryName/>\n"+
                        "<VAT-ID/>\n"+
                        "<Total>"+o.Total+"</Total>\n" +
                        "<Currency>PLN</Currency>\n" +
                        "<ExchangeRate>1</ExchangeRate>\n" +
                        "<SellDate>2017-11-10</SellDate>\n" +
                        "<DeliveryCost>14</DeliveryCost>\n" +
                        "<DeliveryType>Przesyłka kurierska</DeliveryType>\n"+
                        "<PaymentType>"+payment+"</PaymentType>\n"+
                        "<SellerId>14032223</SellerId>\n"+
                        "<positions>\n" +
                        "<position>\n" +
                        "<transactionId>35285349</transactionId>\n" +
                        "<Name/>\n" +
                        "<Quantity>3</Quantity>\n" +
                        "<Price>23.5</Price>\n" +
                        "<OfferName>Wąż asenizacyjny superelastyczny 45mm</OfferName>\n" +
                        "<Signature/>\n" +
                        "</position>\n" +
                        "</positions>\n";
                                
                file += "</transaction>\n";

            }

            file += "</transactions>";
        }



    }
}
