using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public class Furgone : Veicolo
    {
        public Furgone(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo, string image,
            int portata) : base(marca, modello, targa, dtImmatricolazione, prezzo, image)
        {
            Portata = portata;
        }

        public int Portata { get; set; }

        public override string ToString()
        {
            string retVal = $"FURGONE\n{base.ToString()}\nPortata: {Portata}";
            return retVal;
        }

    }
}
