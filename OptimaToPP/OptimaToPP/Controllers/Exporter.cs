using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OptimaToPP.Controllers
{
    public static class Exporter
    {
        public static async Task<int> ToExcel(List<Pack> packs, string fileName)
        {
            return await Task.Run(() =>
            {
                // Load Excel application
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                // Create empty workbook
                excel.Workbooks.Add();

                // Create Worksheet from active sheet
                Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excel.ActiveSheet;

                try
                {
                    workSheet.Cells[1, "A"] = PpColumns._sending_number;
                    workSheet.Cells[1, "B"] = PpColumns._customerName1;
                    workSheet.Cells[1, "C"] = PpColumns._customerName2;
                    workSheet.Cells[1, "D"] = PpColumns._customerStreet;
                    workSheet.Cells[1, "E"] = PpColumns._customerHome1;
                    workSheet.Cells[1, "F"] = PpColumns._customerHome2;
                    workSheet.Cells[1, "G"] = PpColumns._customerPost;
                    workSheet.Cells[1, "H"] = PpColumns._customerCity;
                    workSheet.Cells[1, "I"] = PpColumns._customerCountry;
                    workSheet.Cells[1, "J"] = PpColumns._customerEmail;
                    workSheet.Cells[1, "K"] = PpColumns._customerMobilePhone;
                    workSheet.Cells[1, "L"] = PpColumns._customerPhone;
                    workSheet.Cells[1, "M"] = PpColumns._weight;
                    workSheet.Cells[1, "N"] = PpColumns._cashOnDelivery;
                    workSheet.Cells[1, "O"] = PpColumns._NBR;
                    workSheet.Cells[1, "P"] = PpColumns._tranfserTitle;
                    workSheet.Cells[1, "R"] = PpColumns._comments1;
                    workSheet.Cells[1, "S"] = PpColumns._comments2;
                    workSheet.Cells[1, "T"] = PpColumns._payer;

                    int row = 2;
                    foreach (Pack pack in packs)
                    {
                        string street, city;



                        if (!pack.PostCity.Equals(string.Format(pack.City)) && !String.IsNullOrEmpty(pack.PostCity))
                        {
                            if (!String.IsNullOrEmpty(pack.Street))
                            {
                                street = string.Format("{0}, ul. {1}", pack.City, pack.Street);
                            }
                            else
                            {
                                street = string.Format(pack.City);
                            }
                            city = pack.PostCity;
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

                    MessageBox.Show(string.Format("Pomyślnie zapisano plik: \n{0}", fileName), "Zapisano plik");
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
            });

        }
    }
}
