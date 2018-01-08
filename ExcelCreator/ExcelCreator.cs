using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelCreator
{
    public class ExcelCreator<T>
    {
        String _nombreArchivo = "Libro.xlsx";
        public String NombreArchivo
        {
            get
            {
                return _nombreArchivo.EndsWith(".xlsx") ? _nombreArchivo : _nombreArchivo + ".xlsx";
            }
            set
            {
                _nombreArchivo = value;
            }
        }
        public String Password { get; set; }
        public String Error { get; private set; }
        public String MensajeException { get; private set; }


        List<ExcelHoja<T>> _hojas = new List<ExcelHoja<T>>();
        ExcelConfig _config;
        ExcelMensajes _mensajes;

        public bool AddHoja(ExcelHoja<T> hoja)
        {
            try
            {
                LimpiarMensajes();
                if (hoja == null) return false;

                if (String.IsNullOrEmpty(hoja.Nombre)) hoja.Nombre = _config.NombreHoja;

                if (hoja.Contenido == null)
                {
                    if (hoja.Datasource == null)
                    {
                        Error = _mensajes.NoContenidoNoDatasource;
                        return false;
                    }
                    hoja.Contenido = hoja.ParseMethod == null ? _parseMethodDefault(hoja.Datasource) : hoja.ParseMethod(hoja.Datasource);
                }
                if (hoja.Contenido == null)
                {
                    Error = _mensajes.NoParser;
                    return false;
                }
                if (hoja.Titulos == null)
                {
                    hoja.Titulos = _getTitulos();
                }
                _hojas.Add(hoja);
                return true;
            }
            catch (Exception e)
            {
                Error = _mensajes.FalloAgregarHoja;
                MensajeException = e.Message;
                return false;
            }
        }
        List<ExcelCelda> _getTitulos()
        {
            var titulos = new List<ExcelCelda>();
            foreach (var prop in typeof(T).GetProperties() )
            {
                titulos.Add(new ExcelCelda { Valor = prop.Name });
            }
            return titulos;
        }

        /// <summary>
        /// si no se indicado metodo parser se ejecuta este, haciendo un ToString() de cada propiedad
        /// </summary>
        /// <param name="datasource">datasorce a parsear</param>
        /// <returns>hoja de datos</returns>
        List<List<ExcelCelda>> _parseMethodDefault(List<T> datasource)
        {
            var hoja = new List<List<ExcelCelda>>();

            var t = datasource.FirstOrDefault();
            Type tipo = t.GetType();
            IList<PropertyInfo> props = new List<PropertyInfo>(tipo.GetProperties());

            foreach (var data in datasource)
            {
                var fila = new List<ExcelCelda>();
                foreach (PropertyInfo prop in props)
                {
                    fila.Add(new ExcelCelda
                    {
                        Valor = prop.GetValue(t, null).ToString()
                    });
                }
                hoja.Add(fila);
            }

            return hoja;
        }

        public ExcelCreator()
        {
            try
            {
                LimpiarMensajes();
                using (StreamReader r = new StreamReader(ExcelConfig.RutaJsonConfig))
                {
                    string json = r.ReadToEnd();
                    _config = JsonConvert.DeserializeObject<ExcelConfig>(json);
                }
                if (_config.Lenguaje != null && _config.RutaLenguajes != null)
                {
                    using (StreamReader r = new StreamReader(_config.RutaLenguajes + _config.Lenguaje + ".json"))
                    {
                        string json = r.ReadToEnd();
                        _mensajes = JsonConvert.DeserializeObject<ExcelMensajes>(json);
                    }
                }
            }
            catch (Exception e)
            {
                Error = _mensajes.NoConfig;
                MensajeException = e.Message;
            }

        }

        /// <summary>
        /// limpia mensajes de error. Se debe invocar en todos los metodos publicos
        /// </summary>
        void LimpiarMensajes()
        {
            Error = null;
            MensajeException = null;
        }

        /// <summary>
        /// genera el excel correspondiente en base a contenido
        /// </summary>
        /// <returns>array de bytes del excel</returns>
        public byte[] GenerarExcel()
        {
            try
            {
                LimpiarMensajes();
                var excel = _createExcel();
                var ar = excel.GetAsByteArray();
                if (ar == null)
                {
                    Error = _mensajes.FalloGenerarExcel;
                }
                return ar;
            }
            catch (Exception e)
            {
                Error = _mensajes.FalloGenerarExcel;
                MensajeException = e.Message;
                return null;
            }
        }

        /// <summary>
        /// crear el archivo excel correspondiente en base a contenido, se accedera mediante NombreArchivo
        /// </summary>
        /// <returns>bolean con resultado de operacion</returns>
        public bool CrearArchivoExcel()
        {
            try
            {
                LimpiarMensajes();

                if (String.IsNullOrEmpty(NombreArchivo))
                {
                    Error = _mensajes.NoNombreArchivo;
                    return false;
                }
                var excel = _createExcel();
                var file = new FileInfo(NombreArchivo);
                if (String.IsNullOrEmpty(Password))
                {
                    excel.SaveAs(file);
                }
                else
                {
                    excel.SaveAs(file, Password);
                }

                return true;
            }
            catch (Exception e)
            {
                Error = e.Message;
                return false;
            }
        }

        ExcelPackage _createExcel()
        {
            ExcelPackage excel = new ExcelPackage();
            foreach (var hoja in _hojas)
            {
                int fila = 1, columna = 1;
                var sheet = excel.Workbook.Worksheets.Add(hoja.Nombre);

                foreach (var titulo in hoja.Titulos)
                {
                    sheet.Cells[fila, columna].Value = titulo.Valor;
                    sheet.Cells[fila, columna].Style.Font.Bold = titulo.Negrita;
                    sheet.Cells[fila, columna].Style.Font.UnderLine = titulo.Subrayado;
                    sheet.Cells[fila, columna].Style.Font.Strike = titulo.Tachado;
                    if (titulo.Borde)
                    {
                        sheet.Cells[fila, columna].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    columna++;
                }
                if (hoja.Titulos.Count > 0)
                {
                    fila++;
                    columna = 1;
                }

                foreach (var registro in hoja.Contenido)
                {
                    foreach (var celda in registro)
                    {
                        sheet.Cells[fila, columna].Value = celda.Valor;
                        sheet.Cells[fila, columna].Style.Font.Bold = celda.Negrita;
                        sheet.Cells[fila, columna].Style.Font.UnderLine = celda.Subrayado;
                        sheet.Cells[fila, columna].Style.Font.Strike = celda.Tachado;
                        if (celda.Borde)
                        {
                            sheet.Cells[fila, columna].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        }
                        columna++;
                    }
                    fila++; columna = 1;
                }
            }
            return excel;
        }
    }
}