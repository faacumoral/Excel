using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelReader
{
    public class ExcelReader
    {
        ExcelPackage _excel = new ExcelPackage();

        /// <summary>
        /// si se agrega una entidad por mas que haya error en algun parseo
        /// </summary>
        Boolean _agregarTodos = true;
        public Boolean AgregarTodos { get { return _agregarTodos;  } set { _agregarTodos = value; } }

        /// <summary>
        /// los errores de parseo de los datos
        /// </summary>
        public List<ErrorParseo> ErroresParseo { get; set; }

        /// <summary>
        /// constructor por default con bytes array
        /// </summary>
        /// <param name="file"></param>
        public ExcelReader(byte[] file)
        {
            _excel = new ExcelPackage(new MemoryStream(file));
        }

        public List<T> LeerHojaPorNombre<T>(string nombreHoja) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[nombreHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El nombre de la hoja no es válido");
            }
            return _leerHoja<T>(hoja);
        }

        public List<T> LeerHojaPorIndice<T>(int numeroHoja) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[numeroHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El numero de la hoja no es válido");
            }
            return _leerHoja<T>(hoja);
        }

        public List<T> LeerHoja<T>() where T : new()
        {
            return _leerHoja<T>(_excel.Workbook.Worksheets[1]);
        }

        List<T> _leerHoja<T>(ExcelWorksheet hoja) where T : new()
        {
            ErroresParseo = new List<ErrorParseo>();
            var result = new List<T>();
            int columna = 1, fila = 1;

            try
            {
                var T_properties = typeof(T).GetProperties().Where(p => p.GetSetMethod() != null).ToDictionary(p => p.Name, p => p.GetSetMethod());
                var hoja_properties = new Dictionary<int, string>();

                for (fila = hoja.Dimension.Start.Column;
                                    fila <= hoja.Dimension.End.Column;
                                    fila++)
                {
                    hoja_properties.Add(fila, hoja.Cells[columna, fila].Text);
                }
                fila++;

                for (columna = hoja.Dimension.Start.Row + 1;
                     columna <= hoja.Dimension.End.Row;
                     columna++)
                {
                    var t = new T();
                    var errorParseo = false;
                    for (fila = hoja.Dimension.Start.Column;
                            fila <= hoja.Dimension.End.Column;
                            fila++)
                    {
                        // busco el nombre de la propiedad segun el titulo
                        var propiedad = hoja_properties[fila];

                        // busco la propiedad del metodo
                        if (T_properties.ContainsKey(propiedad))
                        {
                            var setMethod = T_properties[propiedad];
                            var parametro = setMethod.GetParameters().FirstOrDefault().ParameterType;
                            try
                            {
                                // intento parsear valor
                                var valor = Convert.ChangeType(hoja.Cells[columna, fila].Value, parametro);
                                setMethod.Invoke(t, new object[] { valor });
                            }
                            catch (FormatException)
                            {
                                errorParseo = true;
                                // no se pudo parsear
                                ErroresParseo.Add( new ErrorParseo {
                                    Descripcion = "No se pudo parsear '" + hoja.Cells[columna, fila].Value.ToString() + "' como " + parametro.FullName,
                                    Fila = fila,
                                    Columna = columna
                                });
                            }

                        }
                    }
                    if (_agregarTodos || !errorParseo)
                    {
                        result.Add(t);
                    }
                    
                }

                return result;

            }
            catch (Exception)
            {
                return null ;
            }

        }
    }
}
