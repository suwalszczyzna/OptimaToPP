using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;

namespace OptimaToPP
{
    sealed class PackMap : ClassMap<Pack>
    {
        public PackMap()
        {

            AutoMap();
            Map(m => m.DocNumber).Name("Numer dokumentu");
            Map(m => m.Name).Name("Odbiorca");
            Map(m => m.Street).Name("ODBUlica");
            Map(m => m.PostCity).Name("ODBPoczta");
            Map(m => m.NumberHome1).Name("ODBNrDomu");
            Map(m => m.NumberHome2).Name("ODBNrLokalu");
            Map(m => m.ZipCode).Name("ODBKod");
            Map(m => m.City).Name("ODBMiasto");
            Map(m => m.Total).Name("Brutto");
            Map(m => m.Payment).Name("Forma płatności");
            Map(m => m.Email).Name("ODBEmail");
            Map(m => m.Phone).Name("ODBTelefonKom");

        }
    }
}
