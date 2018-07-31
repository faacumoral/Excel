using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCreator
{
    public class ExcelHelpers
    {
        public static List<ExcelCelda> GetCeldas(params object[] valores)
        {
            return valores.Select(v => new ExcelCelda { Valor = v.ToString() }).ToList();
        }
    }
}
