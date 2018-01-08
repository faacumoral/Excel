using System;
using System.Collections.Generic;

namespace ExcelCreator
{
    public class ExcelHoja<T>
    {
        List<ExcelCelda> _titulos = new List<ExcelCelda>();
        Boolean _mostrarTitulos = true;

        public Boolean MostrarTitulos { get { return _mostrarTitulos; } set { _mostrarTitulos = value; } }
        public List<ExcelCelda> Titulos { get { return _titulos == null ? new List<ExcelCelda>() : _titulos; } set { _titulos = value; } }
        public String Nombre { get; set; }
        public List<List<ExcelCelda>> Contenido { get; set; }
        public Func<List<T>, List<List<ExcelCelda>>> ParseMethod { get; set; }
        public List<T> Datasource { get; set; }
    }
}