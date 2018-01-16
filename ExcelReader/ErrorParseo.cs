using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class ErrorParseo
    {
        const string _letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public int Fila { get; set; }
        public int Columna { get; set; }
        public String Descripcion { get; set; }
        public string CeldaExcel {
            get
            {
                string celda = "";
                if (Columna >= _letras.Length)
                    celda += _letras[Columna / _letras.Length - 1];

                return _letras[Columna % _letras.Length - 1].ToString() + Fila.ToString();
            }
        }
        public object ValorLeido { get; set; }

        public String TipoEsperado { get; set; }
    }
}
