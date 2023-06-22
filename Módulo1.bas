Attribute VB_Name = "Módulo1"
Sub minili()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
ws.Activate
'-----------------------Impresion inicial----------------------
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly change"
Range("K1").Value = "Percent change"
Range("L1").Value = "Total Stock Volume"
Range("N2").Value = "Gratest % Increase"
Range("N3").Value = "Gratest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'-----------------------Declaracion de las variables---------------
Dim i, k, l1, l2, sum, gp, gd, gv, j, l As Double

Dim closev, openv, yc, pc As Double
Dim ian, iact As Double
'Obtiene el numero de filas
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'----------------------Inicializa las variables--------------------
'inicia en 2 j para ir imprimiendo los valores sin saltos
sum = 0
sum2 = 0
j = 2
ian = 2
        For i = 2 To LastRow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    'guarda el valor donde encontró la diferencia
                    iact = i
                    'imprime ticker value
                    Cells(j, 9).Value = Cells(i, 1).Value
                    '-----------volumen de compra------------
                    'Va guardando el volumen de compra
                    sum = sum + Cells(i, 7).Value
                    'Imprime el volumen
                    Cells(j, 12).Value = sum
                    'Reinicia el valor de sum
                    sum = 0
                    '-------------------Valor de apertura------
                    'Guarda  el valor de open
                    openv = Cells(ian, 3).Value
                    'imprime open
                    'Cells(j, 10).Value = openv
                    'Suma a i anterior 1 para obtener valor de open
                    ian = i + 1
                    '-------------------Valor de cierre---------
                    'Guarda close
                    closev = Cells(iact, 6).Value
                    '-----------------Year change--------------
                    yc = closev - openv
                    'Imprime el valor de yc
                    Cells(j, 10).Value = yc
                    'Cambio de color verde positivo rojo negativo
                    If yc > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                    Else
                    Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    '-----------Calculo del porcentaje-------------
                    pc = yc / openv
                    'Imprime el valor de pc con color
                    ws.Cells(j, 11).Value = Format(pc, "Percent")
                    
                    
                    'Imprime el calor de close
                    'Cells(j, 11).Value = closev
                    'Va sumando a j que es donde se imprime
                    j = j + 1
                   
                Else
                    'Suma el valor de volumen
                    sum = sum + Cells(i, 7).Value
                End If
        Next i
        '---------------------Inicio GP GD GV------------------------
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox (LastRow2)
        'initialize variables
        gp = Range("K2").Value
        gd = Range("K2").Value
        gv = Range("L2").Value
        
        '---------------------GP------------------------
        For k = 2 To LastRow2
            If Cells(k + 1, 11).Value > gp Then
            gp = Cells(k + 1, 11).Value
            'MsgBox (k)
            l = k + 1
            Else
            gp = gp
            End If
        Next k
         'Imprime el mas alto
         'Cells(2, 16).Value = gp
         'MsgBox (gp)
         '---------------------GD------------------------
        For k = 2 To LastRow2
            If Cells(k + 1, 11).Value < gd Then
            gd = Cells(k + 1, 11).Value
            'MsgBox (k)
            l1 = k + 1
            Else
            gd = gd
            End If
        Next k
        '---------------------GV------------------------
        For k = 2 To LastRow2
            If Cells(k + 1, 12).Value > gv Then
            gv = Cells(k + 1, 12).Value
            'MsgBox (k)
            l2 = k + 1
            Else
            gv = gv
            End If
        Next k
        '---------------------Impresion GPGVGD------------------
         'Imprime el mas alto
         ws.Cells(2, 16).Value = Format(gp, "Percent")
         'Imprime el mas bajo
         ws.Cells(3, 16).Value = Format(gd, "Percent")
         'Imprime el volumen mayor
         ws.Cells(4, 16).Value = Format(gv, "Scientific")
         'MsgBox (gv)
       
         'Imprime el ticker del mas alto
         Cells(2, 15).Value = Cells(l, 9).Value
         Cells(3, 15).Value = Cells(l1, 9).Value
         Cells(4, 15).Value = Cells(l2, 9).Value
      
                
   
Next ws
End Sub



