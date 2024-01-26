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
        /// Load a list of Veicolo from a JSON file
        /// </summary>
        /// <param name="fileName">The file path of the data (parco-auto.json as default)</param>
        /// <returns>A List of Veicolo</returns>
        /// <exception>Something goes wrong reading the data</exception>
        public static List<Veicolo> CaricaDati(string fileName = defaultFileName)
        {
            string jsonString = File.ReadAllText(fileName);
            return JsonSerializer.Deserialize<List<Veicolo>>(jsonString);
        }

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
                var options = new JsonSerializerOptions { WriteIndented = true };
                string jsonString = JsonSerializer.Serialize<List<Veicolo>>(dati, options);
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
