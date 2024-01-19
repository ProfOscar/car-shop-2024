using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public enum TipoAlimentazione
    {
        Undefined,
        Benzina,
        Diesel,
        GPL,
        Metano,
        PlugInHybrid,
        Elettrica
    }

    public class Auto : Veicolo
    {
        public Auto(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo,
            TipoAlimentazione alimentazione, int numPorte, bool isTettoPanoramico):base(marca, modello, targa, dtImmatricolazione, prezzo)
        {
            Alimentazione = alimentazione;
            NumPorte = numPorte;
            IsTettoPanoramico = isTettoPanoramico;
        }

        public TipoAlimentazione Alimentazione { get; set; }
        public int NumPorte { get; set; }
        public bool IsTettoPanoramico { get; set; }

        public override string ToString()
        {
            string retVal = $"AUTO\n{base.ToString()}";
            retVal += $"\n{Alimentazione} {NumPorte} Porte";
            if (IsTettoPanoramico) retVal += "\nCon Tetto Panoramico!";
            return retVal;
        }
    }
}
