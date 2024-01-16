using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public class Auto : Veicolo
    {
        public Auto(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo,
            int numPorte, bool isTettoPanoramico):base(marca, modello, targa, dtImmatricolazione, prezzo)
        {
            NumPorte = numPorte;
            IsTettoPanoramico = isTettoPanoramico;
        }

        public int NumPorte { get; set; }
        public bool IsTettoPanoramico { get; set; }

        public override string ToString()
        {
            string retVal = $"AUTO\n{base.ToString()}\nN.Porte: {NumPorte}";
            if (IsTettoPanoramico) retVal += "\nCon Tetto Panoramico!";
            return retVal;
        }
    }
}
