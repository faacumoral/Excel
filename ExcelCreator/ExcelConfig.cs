using System;

namespace ExcelCreator
{
    public class ExcelConfig
    {
        string _rutaLenguajes;
        public string NombreArchivo { get; set; }
        public string Lenguaje { get; set; }
        public string RutaLenguajes {
            get
            { return new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)).LocalPath + _rutaLenguajes; }
            set
            { _rutaLenguajes = value; }
        }
        public string NombreHoja { get; set; }

        public static string RutaJsonConfig = new Uri(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)).LocalPath + "/JSON/ExcelConfig.json";
    }
}