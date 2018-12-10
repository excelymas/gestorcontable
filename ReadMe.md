Gestor Contable en Excel
========================
[Suscríbete a mi canal Excel & Más en YouTube, donde encontrarás muchísimo contenido gratuito](https://www.youtube.com/user/ottojaviergonzalez)


>Si quieres ver a detalle la explicación del Inicio del Proyecto: [Ver Vídeo Sesión 1](https://youtu.be/AejTkvK97pY)


Este es un proyecto desarrollado en Excel mediante Visual Basic para Aplicaciones y guardado en formato binario. Está orientado a la creación de macros simples y amigables que permiten a los profesionales del área contable a ser más productivos en sus negocios.

La versión 0.1.0 de este proyecto, contiene las dos primeras macros que facilitarán las tareas de generar de forma automática un Libro Mayor y un Balance de Comprobación.

Una función que juega un papel muy importante en el proyecto es la función nReg, que permite identificar la fila inicial y la última fila en un listado, y que además puede identificar la celda vacía después del último registro en un listado, lo que facilita agregar nuevos registros a una base de datos si así se requiere.


>A continuación, te presento la función nReg
```vb
Public Function nReg(Hoja As Worksheet, nFila As Long, nColumna As Long) As Long
    Do Until IsEmpty(Hoja.Cells(nFila, nColumna))
        nFila = nFila + 1
    Loop
    nReg = nFila
End Function
```


>La macro que nos permite generar de forma automática un Libro Mayor es la siguiente
```vb
Sub LibroMayor() 'MACRO PARA GENERAR EL LIBRO MAYOR
Dim Encontrado As Boolean
Dim Encabezado As Boolean
Dim i As Long
Dim ccFila As Long, ccFinal As Long
Dim ldFila As Long, ldFinal As Long
Dim vDebe As Double, vHaber As Double

ccFinal = nReg(Hoja2, 2, 1) - 1 ' Catálogo de Cuentas
ldFinal = nReg(Hoja3, 2, 1) - 1 ' Libro Diario
i = 1

With Hoja4 ' LIBRO MAYOR
.Activate
.Range("A:G").Clear

    For ccFila = 2 To ccFinal
        If Len(Hoja2.Cells(ccFila, 1)) = 3 Then
                vDebe = 0
                vHaber = 0
                Encabezado = True
                Encontrado = False
            For ldFila = 2 To ldFinal
                    If Hoja2.Cells(ccFila, 1) = Mid(Hoja3.Cells(ldFila, 4), 1, 3) Then
                            Encontrado = True
                                    If Encabezado = True Then
                                        'ENCABEZADO
                                        .Cells(i, 1) = "CUENTA"
                                        .Cells(i, 2) = "NOMBRE DE LA CUENTA"
                                        .Cells(i, 3) = "ASIENTO #"
                                        .Cells(i, 4) = "FECHA"
                                        .Cells(i, 5) = "DEBE"
                                        .Cells(i, 6) = "HABER"
                                        .Cells(i, 7) = "SALDO"
                                        'FORMATO
                                        .Range(Cells(i, 1), .Cells(i, 7)).HorizontalAlignment = xlCenter
                                        .Range(.Cells(i, 1), .Cells(i, 7)).Interior.Color = RGB(190, 190, 190)
                                        .Range(.Cells(i, 1), .Cells(i, 7)).Font.Color = RGB(255, 255, 255)
                                        .Range(.Cells(i, 1), .Cells(i, 7)).Font.Bold = True
                                        i = i + 1
                                        Encabezado = False
                                    End If
                            'Registros
                            .Cells(i, 1) = Hoja2.Cells(ccFila, 1) 'No. Cuenta
                            .Cells(i, 2) = Hoja2.Cells(ccFila, 2) 'Nombre Cuenta
                            .Cells(i, 3) = Hoja3.Cells(ldFila, 1) 'No. Partida o Asiento
                            .Cells(i, 4) = Hoja3.Cells(ldFila, 2) 'Fecha
                            .Cells(i, 4).NumberFormat = "dd/mm/yyyy"
                            'SUMANDO DEBE Y HABER
                            .Cells(i, 5) = Hoja3.Cells(ldFila, 6) 'DEBE
                            .Cells(i, 6) = Hoja3.Cells(ldFila, 7) 'HABER
                            'ASIGNAR SUMAS AL DEBE Y EL HABER
                            vDebe = vDebe + Hoja3.Cells(ldFila, 6) 'DEBE
                            vHaber = vHaber + Hoja3.Cells(ldFila, 7) 'HABER
                             'SALDO
                            .Cells(i, 7) = vDebe - vHaber 'SALDO
                            'FORMATO MONEDA
                            .Range(.Cells(i, 5), .Cells(i, 7)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                            
                            i = i + 1
                    End If
            Next ldFila
                    If Encontrado = True Then
                        i = i + 1
                        'TOTALES
                        .Cells(i, 4) = "Total: "
                        .Cells(i, 5) = vDebe 'Total DEBE
                        .Cells(i, 6) = vHaber 'Total HABER
                        .Cells(i, 7) = vDebe - vHaber 'Total SALDO
                        'FORMATO MONEDA Y NEGRITA
                        .Range(.Cells(i, 5), .Cells(i, 7)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        .Range(.Cells(i, 4), .Cells(i, 7)).Font.Bold = True
                        i = i + 2
                    End If
        End If
    Next ccFila
    
    .Range("A:G").EntireColumn.AutoFit
End With

End Sub
```

La misma macro utilizada para la construcción del Libro Mayor, fue modificada y adaptada para crear la automatización del Balance de Comprobación.

```vb
Sub BalanceComprobacion() 'MACRO PARA GENERAR EL BALANCE DE COMPROBACIÓN
Dim Encontrado As Boolean
Dim Encabezado As Boolean
Dim i As Long
Dim ccFila As Long, ccFinal As Long
Dim ldFila As Long, ldFinal As Long
Dim vDebe As Double, vHaber As Double
Dim totDebe As Double, totHaber As Double
Dim totSaldoDeudor As Double, totSaldoAcreedor As Double

ccFinal = nReg(Hoja2, 2, 1) - 1 ' Catálogo de Cuentas
ldFinal = nReg(Hoja3, 2, 1) - 1 ' Libro Diario
i = 1

With Hoja5 ' BALANCE DE COMPROBACIÓN
.Activate
.Range("A:F").Clear
    
Encabezado = True
    
    For ccFila = 2 To ccFinal
        If Len(Hoja2.Cells(ccFila, 1)) = 3 Then
                    vDebe = 0
                    vHaber = 0
                    Encontrado = False
                For ldFila = 2 To ldFinal
                    If Hoja2.Cells(ccFila, 1) = Mid(Hoja3.Cells(ldFila, 4), 1, 3) Then
                        Encontrado = True
                            If Encabezado = True Then
                                'ENCABEZADO
                                .Cells(i, 1) = "CUENTA"
                                .Cells(i, 2) = "NOMBRE DE LA CUENTA"
                                .Cells(i, 3) = "DEBE"
                                .Cells(i, 4) = "HABER"
                                .Cells(i, 5) = "SALDO DEUDOR"
                                .Cells(i, 6) = "SALDO ACREEDOR"
                                'FORMATO
                                .Range(Cells(i, 1), .Cells(i, 6)).HorizontalAlignment = xlCenter
                                .Range(.Cells(i, 1), .Cells(i, 6)).Interior.Color = RGB(190, 190, 90)
                                .Range(.Cells(i, 1), .Cells(i, 6)).Font.Color = RGB(255, 255, 255)
                                .Range(.Cells(i, 1), .Cells(i, 6)).Font.Bold = True
                                i = i + 1
                                Encabezado = False
                            End If
                        'Suma de sub totales y totales del DEBE y el HABER
                        vDebe = vDebe + Hoja3.Cells(ldFila, 6) 'Sub Total Debe
                        vHaber = vHaber + Hoja3.Cells(ldFila, 7) 'Sub Total Haber
                        
                        totDebe = totDebe + Hoja3.Cells(ldFila, 6) 'Total Debe
                        totHaber = totHaber + Hoja3.Cells(ldFila, 7) 'Total Haber
                    End If
                Next ldFila
                            If Encontrado = True Then
                                'REGISTROS
                                .Cells(i, 1) = Hoja2.Cells(ccFila, 1) 'No. Cuenta
                                .Cells(i, 2) = Hoja2.Cells(ccFila, 2) 'Nombre Cuenta
                                .Cells(i, 3) = vDebe 'Sub Total Debe
                                .Cells(i, 4) = vHaber 'Sub Total Haber
                                
                                 
                                        'SALDO DEUDOR Y SALDO ACREEDOR
                                        If vDebe > vHaber Then
                                                'VALORES
                                                .Cells(i, 5) = vDebe - vHaber 'Saldo Deudor
                                                .Cells(i, 6) = 0
                                                totSaldoDeudor = totSaldoDeudor + .Cells(i, 5)
                                            Else
                                                'VALORES
                                                .Cells(i, 5) = 0
                                                .Cells(i, 6) = vHaber - vDebe 'Saldo Acreedor
                                                totSaldoAcreedor = totSaldoAcreedor + .Cells(i, 6)
                                        End If
                                        
                                'FORMATO MONEDA
                                .Range(.Cells(i, 3), .Cells(i, 6)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                                i = i + 1
                            End If
        End If
        
        If ccFila = ccFinal Then
            i = i + 1
            'TOTALES
            .Cells(i, 2) = "Totales ------------------------------------------------------>"
            .Cells(i, 3) = totDebe 'Total Debe
            .Cells(i, 4) = totHaber 'Total Haber
            .Cells(i, 5) = totSaldoDeudor 'Total Saldo Deudor
            .Cells(i, 6) = totSaldoAcreedor 'Total Saldo Acredor
            'FORMATO MONEDA Y NEGRITA
            .Range(.Cells(i, 3), .Cells(i, 6)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Range(.Cells(i, 2), .Cells(i, 6)).Font.Bold = True
        End If
        
    Next ccFila
    
    .Range("A:F").EntireColumn.AutoFit
End With

End Sub
```

