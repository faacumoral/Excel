# Excel

ExcelCreator: para crear un excel se deben seguir los siguientes pasos:
1) Crear objeto excel: var excel = new Excel<T>();
   <T>: tipo de datos que contendrán las hojas (esto es únicamente para cuando se setea el contenido de la hoja mediante una lista y un método parseador de contenido)
     .NombreArchivo (String): nombre del archivo 
     .Password (String): opcional, contraseña para colocarle al archivo
     .Error (String): mensaje de error arrojado mediante algun proceso (friendly-user)
     .MensajeException: mensaje de la excepción para tratado por el desarrollador
     
2) Agregar hojas: excel.AddHoja(new ExcelHoja<T>());
  donde se pueden definir:
    .Titulos (List<ExcelCelda>): encabezados del excel
    
    .Nombre (String): nombre de la hoja
    .Contenido (List<List<ExcelCelda>>): matriz (fila x columna) con el contenido de la hoja
    .Datasource (List<T>): lista de entidades a guardar en la hoja (donde cada entidad va a representar una fila
    .ParseMethod (Func<List<T>, List<List<ExcelCelda>>>): método que parsear el datasource a un List<List<ExcelCelda>> para escribir en la hoja
  Para setear el contenido de la hoja se puede o bien definir el contenido manualmente (mediante la propiedad .Contenido) o mediante  .Datasource y .ParseMethod, pero setear el contenido de la hoja de ambas formas es innecesario.

  3) Generar archivo: una vez seteada toda la información, se deberá crear el archivo, para lo cual se puede generar un byte array (excel.GenerarExcel(), retorna null si falló) o escribir el archivo en la ubicación raíz (exce.CrearArchivoExcel(), que retorna booleano con resultado de la operación)
  
  Asimismo, a una celda (ExcelCelda) se le pueden definir los siguientes valores:
  .Valor (String): contenido de la celda
  .Negrita (Boolean)
  .Subrayado (Boolean)
  .Tachado (Boolean)
  .Borde (Boolean)
  
  Por último, se pueden modificar ciertas configuraciones por default:
  JSON/ExcelConfig.json:
    NombreArchivo: nombre para el archivo en caso de no inidicarle ninguna
    NombreHoja: nombre por default en caso de no inidicar ninguno
    Lenguaje: idioma para los mensajes de errores (debe existir el archivo lenguaje.json dentro de la ruta especificada)
    RutaLenguajes: ruta donde se leeran los lenguajes
    
    
ExcelReader: para leer un excel se deben seguir los siguientes pasos:
1) Crear excel a partir de byte array: var excel = new ExcelReader(byte[] excelBytes);
    donde se puede definir:
        .AgregarTodos (Boolean): si se agrega una entidad en caso de error de parseo de datos
2) Leer datos: los datos se pueden leer de dos formas distintas, por bloque y por hoja. Asimismo, en cada caso se puede indicar si leer horizontal o vertical:
    LeerHoja<T>(string nombreHoja?/int numeroHoja?, bool vertical)
    LeerBloque<T>(string nombreHoja?/int numeroHoja?, bool vertical, int inicioFila, int finFila, int inicioColumna, int finColumna)
    para dichos casos, el nombreHoja y numeroHoja es sobre que hoja se desea leer. Si no se inidican, se tomará la primera. Son opcionales mediante sobrecarga; vertical indica si la lectura se debe hacer vertical u horizontalmente. Asimismo, T es el tipo de datos de la hoja/bloque: se buscará la primer fila/columna y esa indicará los nombres de las propiedades, es decir, si una fila posee el header "Edad", dichos valores se almacenarán en .Edad de T (siempre y cuando tenga método set público).
    Para el segundo tipo, LeerBloque, además se puede inidicar el bloque de fila y columna sobre el cual leer los datos por si una hoja posee varios formatos. Si se indica 0 en alguno de ellos, se tomará el inicio o el final de la hoja según corresponda. Por ejemplo, si se indica 5 como inicio de fila y 0 como fin, se comenzará a leer desde la fila 5 hasta el final de la hoja.
    Cabe destacar que la conversación es autómatica segun el tipo de datos que posea T, esto es, volviendo al ejemplo anterior, si la propiedad .Edad es de tipo int, se intentará parsear solo a int. En caso de fallar, se almacenará en una lista de errores de parseo; además, la conversación a T es flexible, esto es, no es necesario se definan todas la propiedades de T en el excel ni viceversa, se parsearan las que coincidan.
3) Ver errores de parseo (si existieron): mediante excel.ErroresParseo (List<ErrorParseo>) después de la lectura. De los errores podemos observar:
    .Fila (int): fila del errores
    .Columna (int): columna del error
    .Descripcion (String): descripcion de parseo esperado 
    .CeldaExcel (String): nombre de la celda excel (A1, B4, etc)