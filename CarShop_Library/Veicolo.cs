using System;
using System.Text.Json.Serialization;

namespace CarShop_Library
{

    [JsonDerivedType(typeof(Auto), typeDiscriminator: "auto")]
    [JsonDerivedType(typeof(Moto), typeDiscriminator: "moto")]
    [JsonDerivedType(typeof(Furgone), typeDiscriminator: "furg")]
    public abstract class Veicolo
    {
        public Veicolo() { }

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
