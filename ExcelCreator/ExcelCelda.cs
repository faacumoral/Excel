using System;

namespace ExcelCreator
{
    public class ExcelCelda
    {
        public String Valor { get; set; }
        Boolean _negrita = false;
        Boolean _subrayado = false;
        Boolean _tachado = false;
        Boolean _borde = false;

        public Boolean Negrita { get { return _negrita; } set { _negrita = value; } }
        public Boolean Subrayado { get { return _subrayado; } set { _subrayado = value; } }
        public Boolean Tachado { get { return _tachado; } set { _tachado = value; } }
        public Boolean Borde { get { return _borde; } set { _borde = value; } }
    }
}
