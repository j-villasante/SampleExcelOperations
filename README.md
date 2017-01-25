# SampleExcelOperations

Ambas clases son necesarias para el correcto funcionamiento.

## Reader.cs
Esta clase se encarga en este caso de abrir y cerrar el archivo excel. En caso este prototipo se utilice en un proyecto VSTO los metodos *open* y *close* no serían necesarios. Se tendría que utilizar el metodo *setWorkbook*.

## BookEditor.cs
Hereda de Reader.cs. Contiene los metodos para poder agregar y eliminar los nombres de rango. 
