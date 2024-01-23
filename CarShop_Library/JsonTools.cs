using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public class JsonTools
    {
        const string defaultFileName = "parco-auto.json";

        /// <summary>
        /// Save a list of Veicolo in a JSON file
        /// </summary>
        /// <param name="dati">The list of Veicolo to save</param>
        /// <param name="fileName">The file path of the saved data (parco-auto.json as default)</param>
        /// <returns>true if Ok, else false</returns>
        public static bool SalvaDati(List<Veicolo> dati, string fileName = defaultFileName)
        {
            try
            {
                string jsonString = JsonSerializer.Serialize<List<Veicolo>>(dati);
                File.WriteAllText(fileName, jsonString);
                return true;
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
                return false;
            }
        }
    }
}
