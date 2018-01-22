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
        public Boolean AgregarTodos { get { return _agregarTodos; } set { _agregarTodos = value; } }

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

        /// <summary>
        /// leer toda una hoja 
        /// </summary>
        /// <typeparam name="T">tipo de datos de la hoja</typeparam>
        /// <param name="nombreHoja"></param>
        /// <param name="vertical">si lee vertical</param>
        /// <returns></returns>
        public List<T> LeerHoja<T>(string nombreHoja, bool vertical) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[nombreHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El nombre de la hoja no es válido");
            }
            return _leer<T>(hoja, hoja.Dimension.Start.Row, hoja.Dimension.End.Row, hoja.Dimension.Start.Column, hoja.Dimension.End.Column, vertical);
        }

        /// <summary>
        /// leer toda una hoja
        /// </summary>
        /// <typeparam name="T">tipo de datos de la hoja</typeparam>
        /// <param name="numeroHoja"></param>
        /// <param name="vertical">si lee vertical</param>
        /// <returns></returns>
        public List<T> LeerHoja<T>(int numeroHoja, bool vertical) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[numeroHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El numero de la hoja no es válido");
            }
            return _leer<T>(hoja, hoja.Dimension.Start.Row, hoja.Dimension.End.Row, hoja.Dimension.Start.Column, hoja.Dimension.End.Column, vertical);
        }

        /// <summary>
        /// leer toda una hoja 
        /// </summary>
        /// <typeparam name="T">tipo de datos de la hoja</typeparam>
        /// <param name="vertical">si lee vertical</param>
        /// <returns></returns>
        public List<T> LeerHoja<T>(bool vertical) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[1];
            return _leer<T>(hoja, hoja.Dimension.Start.Row, hoja.Dimension.End.Row, hoja.Dimension.Start.Column, hoja.Dimension.End.Column, vertical);
        }


        /// <summary>
        /// leer bloque de celdas; si se pasa 0 en inicio/fin columna/fila se toma el inicio/fin
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="nombreHoja"></param>
        /// <param name="vertical"></param>
        /// <param name="inicioFila"></param>
        /// <param name="finFila"></param>
        /// <param name="inicioColumna"></param>
        /// <param name="finColumna"></param>
        /// <param name=""></param>
        /// <returns></returns>
        public List<T> LeerBloque<T>(string nombreHoja, bool vertical, int inicioFila, int finFila, int inicioColumna, int finColumna) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[nombreHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El nombre de la hoja no es válido");
            }
            inicioFila = inicioFila == 0 ? hoja.Dimension.Start.Row : inicioFila;
            finFila = finFila == 0 ? hoja.Dimension.End.Row : finFila;
            inicioColumna = inicioColumna == 0 ? hoja.Dimension.Start.Column : inicioColumna;
            finColumna = finColumna == 0 ? hoja.Dimension.End.Column : finColumna;

            return _leer<T>(hoja, inicioFila, finFila, inicioColumna, finColumna, vertical);
        }

        /// <summary>
        /// leer bloque de celdas; si se pasa 0 en inicio/fin columna/fila se toma el inicio/fin
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="numeroHoja"></param>
        /// <param name="vertical"></param>
        /// <param name="inicioFila"></param>
        /// <param name="finFila"></param>
        /// <param name="inicioColumna"></param>
        /// <param name="finColumna"></param>
        /// <param name=""></param>
        /// <returns></returns>
        public List<T> LeerBloque<T>(int numeroHoja, bool vertical, int inicioFila, int finFila, int inicioColumna, int finColumna) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[numeroHoja];
            if (hoja == null)
            {
                throw new ArgumentException("El numero de la hoja no es válido");
            }
            inicioFila = inicioFila == 0 ? hoja.Dimension.Start.Row : inicioFila;
            finFila = finFila == 0 ? hoja.Dimension.End.Row : finFila;
            inicioColumna = inicioColumna == 0 ? hoja.Dimension.Start.Column : inicioColumna;
            finColumna = finColumna == 0 ? hoja.Dimension.End.Column : finColumna;

            return _leer<T>(hoja, inicioFila, finFila, inicioColumna, finColumna, vertical);
        }

        /// <summary>
        /// leer bloque de celdas; si se pasa 0 en inicio/fin columna/fila se toma el inicio/fin
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="vertical"></param>
        /// <param name="inicioFila"></param>
        /// <param name="finFila"></param>
        /// <param name="inicioColumna"></param>
        /// <param name="finColumna"></param>
        /// <param name=""></param>
        /// <returns></returns>
        public List<T> LeerBloque<T>(bool vertical, int inicioFila, int finFila, int inicioColumna, int finColumna) where T : new()
        {
            var hoja = _excel.Workbook.Worksheets[1];
            inicioFila = inicioFila == 0 ? hoja.Dimension.Start.Row : inicioFila;
            finFila = finFila == 0 ? hoja.Dimension.End.Row : finFila;
            inicioColumna = inicioColumna == 0 ? hoja.Dimension.Start.Column : inicioColumna;
            finColumna = finColumna == 0 ? hoja.Dimension.End.Column : finColumna;

            return _leer<T>(hoja, inicioFila, finFila, inicioColumna, finColumna, vertical);
        }
        

        List<T> _leer<T>(ExcelWorksheet hoja, int inicioFila, int finFila, int inicioColumna, int finColumna, bool vertical) where T : new()
        {
            if (vertical)
            {
                return _leerVertical<T>(hoja, inicioFila, finFila, inicioColumna, finColumna);
            }
            else
            {
                return _leerHorizontal<T>(hoja, inicioFila, finFila, inicioColumna, finColumna);

            }
        }

        List<T> _leerVertical<T>(ExcelWorksheet hoja, int inicioFila, int finFila, int inicioColumna, int finColumna) where T : new()
        {
            ErroresParseo = new List<ErrorParseo>();
            var result = new List<T>();
            int columna = inicioColumna, fila;

            try
            {
                var T_properties = typeof(T).GetProperties().Where(p => p.GetSetMethod() != null).ToDictionary(p => p.Name, p => p.GetSetMethod());
                var hoja_properties = new Dictionary<int, string>();


                for (fila = inicioFila;
                    fila <= finFila;
                    fila++)
                {
                    hoja_properties.Add(fila, hoja.Cells[fila, columna].Text);
                }
                columna++;

                for (columna = inicioColumna + 1;
                     columna <= finColumna;
                     columna++)
                {
                    var t = new T();
                    var errorParseo = false;
                    bool esFilaVacia = true;
                    for (fila = inicioFila;
                            fila <= finFila;
                            fila++)
                    {
                        // busco el nombre de la propiedad segun el titulo
                        var propiedad = hoja_properties[fila];

                        // busco la propiedad del metodo
                        if (T_properties.ContainsKey(propiedad) && hoja.Cells[fila, columna].Value != null)
                        {
                            esFilaVacia = false;
                            var setMethod = T_properties[propiedad];
                            var parametro = setMethod.GetParameters().FirstOrDefault().ParameterType;
                            try
                            {
                                // intento parsear valor
                                var valor = Convert.ChangeType(hoja.Cells[fila, columna].Value, parametro);
                                setMethod.Invoke(t, new object[] { valor });
                            }
                            catch (FormatException)
                            {
                                errorParseo = true;
                                // no se pudo parsear
                                ErroresParseo.Add(new ErrorParseo
                                {
                                    Fila = fila,
                                    Columna = columna,
                                    ValorLeido = hoja.Cells[fila, columna].Value,
                                    TipoEsperado = parametro.FullName
                                });
                            }
                            catch (InvalidCastException)
                            {
                                errorParseo = true;
                                // no se pudo parsear
                                ErroresParseo.Add(new ErrorParseo
                                {
                                    Fila = fila,
                                    Columna = columna,
                                    ValorLeido = hoja.Cells[fila, columna].Value,
                                    TipoEsperado = parametro.FullName
                                });
                            }

                        }
                    }
                    if ((_agregarTodos || !errorParseo) && !esFilaVacia)
                    {
                        result.Add(t);
                    }

                }

                return result;

            }
            catch (Exception)
            {
                return null;
            }
        }

        List<T> _leerHorizontal<T>(ExcelWorksheet hoja, int inicioFila, int finFila, int inicioColumna, int finColumna) where T : new()
        {
            ErroresParseo = new List<ErrorParseo>();
            var result = new List<T>();
            int columna, fila = inicioFila;

            try
            {
                var T_properties = typeof(T).GetProperties().Where(p => p.GetSetMethod() != null).ToDictionary(p => p.Name, p => p.GetSetMethod());
                var hoja_properties = new Dictionary<int, string>();

                for (columna = inicioColumna;
                    columna <= finColumna;
                    columna++)
                {
                    hoja_properties.Add(columna, hoja.Cells[fila, columna].Text);
                }
                fila++;

                for (fila = inicioFila + 1;
                     fila <= finFila;
                     fila++)
                {
                    var t = new T();
                    bool errorParseo = false;
                    bool esFilaVacia = true;
                    for (columna = inicioColumna;
                            columna <= finColumna;
                            columna++)
                    {
                        // busco el nombre de la propiedad segun el titulo
                        var propiedad = hoja_properties[columna];

                        // busco la propiedad del metodo
                        if (T_properties.ContainsKey(propiedad) && hoja.Cells[fila, columna].Value != null)
                        {
                            esFilaVacia = false;
                            var setMethod = T_properties[propiedad];
                            var parametro = setMethod.GetParameters().FirstOrDefault().ParameterType;
                            try
                            {
                                // intento parsear valor
                                var valor = Convert.ChangeType(hoja.Cells[fila, columna].Value, parametro);
                                setMethod.Invoke(t, new object[] { valor });
                            }
                            catch (FormatException)
                            {
                                errorParseo = true;
                                // no se pudo parsear
                                ErroresParseo.Add(new ErrorParseo
                                {
                                    Fila = fila,
                                    Columna = columna,
                                    ValorLeido = hoja.Cells[fila, columna].Value,
                                    TipoEsperado = parametro.FullName
                                });
                            }
                            catch (InvalidCastException)
                            {
                                errorParseo = true;
                                // no se pudo parsear
                                ErroresParseo.Add(new ErrorParseo
                                {
                                    Fila = fila,
                                    Columna = columna,
                                    ValorLeido = hoja.Cells[fila, columna].Value,
                                    TipoEsperado = parametro.FullName
                                });
                            }
                        }
                    }
                    if ((_agregarTodos || !errorParseo) && !esFilaVacia)
                    {
                        result.Add(t);
                    }

                }

                return result;

            }
            catch (Exception)
            {
                return null;
            }
        }

    }
}
