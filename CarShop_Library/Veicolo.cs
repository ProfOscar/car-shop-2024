using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public abstract class Veicolo
    {
        public Veicolo(string marca, string modello, string targa, DateTime dtImmatricolazione, int prezzo, string image)
        {
            Marca = marca;
            Modello = modello;
            Targa = targa;
            DtImmatricolazione = dtImmatricolazione;
            Prezzo = prezzo;
            Image = image;
        }

        public string Marca { get; set; }
        public string Modello { get; set; }
        public string Targa { get; set; }
        public DateTime DtImmatricolazione { get; set; }
        public int Prezzo { get; set; }
        public string Image { get; set; }

        public override string ToString()
        {
            string retVal = $"{Marca} {Modello} - {Targa}\nImmatricolata: {DtImmatricolazione.ToShortDateString()}\nPrezzo: {Prezzo} EURO";
            return retVal;
        }
    }
}
