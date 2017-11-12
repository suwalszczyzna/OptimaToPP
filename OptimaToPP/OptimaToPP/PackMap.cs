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
            Map(m => m.NrDocument).Name("Numer dokumentu");
            Map(m => m.RecipientName).Name("Kontrahent");
            Map(m => m.RecipientAdress).Name("Ulica");
            Map(m => m.RecipientNoHome).Name("NR_DOMU");
            Map(m => m.RecipientNoHome2).Name("NR_LOKALU");
            Map(m => m.RecipientZIP).Name("KOD_POCZTOWY");
            Map(m => m.RecipientCity).Name("Miasto");
            Map(m => m.Total).Name("Brutto");
            Map(m => m.RecipientPayment).Name("Forma płatności");

                   
        }
    }
}
