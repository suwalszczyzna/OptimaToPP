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
            Map(m => m.RecipientName).Name("Odbiorca");
            Map(m => m.RecipientAdress).Name("ODBUlica");
            Map(m => m.RecipientNoHome).Name("ODBNrDomu");
            Map(m => m.RecipientNoHome2).Name("ODBNrLokalu");
            Map(m => m.RecipientZIP).Name("ODBKod");
            Map(m => m.RecipientCity).Name("ODBMiasto");
            Map(m => m.Total).Name("Brutto");
            Map(m => m.RecipientPayment).Name("Forma płatności");

                   
        }
    }
}
