using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public class Moto : Veicolo
    {
        public Moto(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo,
            int numTempi, bool isBauletto): base(marca, modello, targa, dtImmatricolazione, prezzo)
        {
            NumTempi = numTempi;
            IsBauletto = isBauletto;
        }

        public int NumTempi { get; set; }
        public bool IsBauletto { get; set; }

        public override string ToString()
        {
            string retVal = $"MOTO\n{base.ToString()}\nN.Tempi: {NumTempi}";
            if (IsBauletto) retVal += "\nCon Bauletto!";
            return retVal;
        }
    }
}
