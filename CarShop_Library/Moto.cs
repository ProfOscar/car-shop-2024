using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public enum TipoMoto
    {
        Undefined,
        Cross,
        Enduro,
        Strada,
        Chopper,
        Touring
    }

    public class Moto : Veicolo
    {
        public Moto(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo,
            TipoMoto tipo, int numTempi, bool isBauletto): base(marca, modello, targa, dtImmatricolazione, prezzo)
        {
            Tipo = tipo;
            NumTempi = numTempi;
            IsBauletto = isBauletto;
        }

        public TipoMoto Tipo { get; set; }
        public int NumTempi { get; set; }
        public bool IsBauletto { get; set; }

        public override string ToString()
        {
            string retVal = $"MOTO\n{base.ToString()}";
            retVal += $"\n{Tipo} {NumTempi} Tempi";
            if (IsBauletto) retVal += "\nCon Bauletto!";
            return retVal;
        }
    }
}
