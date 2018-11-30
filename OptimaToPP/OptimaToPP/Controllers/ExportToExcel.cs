using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OptimaToPP.Controllers
{
    public static class ExportToExcel
    {
        public static int Export(List<Pack> packs, string fileName)
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


                    if (!pack.PostCity.Equals(string.Format(pack.City)) && pack.PostCity.Length > 3)
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
                return 0;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Exception",
                    "Błąd podczas zapisu pliku\n" + exception.Message,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
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
