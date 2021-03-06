﻿using CsvHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OptimaToPP
{
    public static class Converter
    {

        public static void XLStoCSV(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {

            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            if (File.Exists(csvOutputFile))
            {
                File.Delete(csvOutputFile);
            }
            // connection string
            var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("{0}", e));

            }
            finally
            {
                // free resources
                cnn.Close();
            }

            // write out CSV data
            using (var wtr = new StreamWriter(csvOutputFile))
            {
                foreach (DataRow row in dt.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (!firstLine) { wtr.Write(";"); } else { firstLine = false; }
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                }
            }
        }

        public static async Task<List<Pack>> CSVtoListOfPacks(string PathToCsv)
        {
            return await Task.Run(() =>
            {
                List<Pack> packs = new List<Pack>();
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
                return packs;
            });
        }
    }
}
