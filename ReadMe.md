Gestor Contable en Excel
========================
Este es un proeycto desarrollado en Excel mediante Visual Basic para Aplicaciones y guardado en formato binario.
Está orientado a la creación de macros simples y amigables que permiten a los profesiones del área contable a ser más productivos en sus negocios.

La versión 0.1.0 contiene las dos primeras macros que facilitarán las tareas de generar de forma automática un Libro Mayor y un Balance de Comprobación.

Una función que juega un papel muy importante en el proyecto, es la función nReg, que pemite identificar la fila inicial y la última fila en un listado, y que además puede identificar la celda vacía después del último registro en un listado, lo que facilita agregar nuevos registros a una base de datos si así se requiere.

A continuación te presento la función nReg

```vb
Public Function nReg(Hoja As Worksheet, nFila As Long, nColumna As Long) As Long
    Do Until IsEmpty(Hoja.Cells(nFila, nColumna))
        nFila = nFila + 1
    Loop
    nReg = nFila
End Function
```

