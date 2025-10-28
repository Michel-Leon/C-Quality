'============inicio del programa inspeccion===================
'=======VARIABLES GLOBALES======
public Palabra As String' SE GUARDA EL TEXTO CORRESPONDIENTE
'----------------------------------------------------------------------
'                   PROCEDIMIENTOS
'----------------------------------------------------------------------
Public Sub buscar_Inspeccion(PalabraBuscada As String,MTM1 As String,MTM2 As String,BRR As String,CP As String,CC As String,CM AS String)
    On Error GoTo ErrorHandler

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim i As Long ' CONTADOR PARA RECCORER LAS FILAS
    Dim filaDestino As Long
    Dim encontrada As Boolean

    '====SELECCIONA LAS HOJAS DE ORIGEN Y DESTINO====
    Set wsOrigen = ThisWorkbook.Worksheets("PROYECTO")
    Set wsDestino = ThisWorkbook.Worksheets("INSPECCION")
    encontrada = False
    
    '====BUSQUEDA DE PALABRA EN ORIGEN (CE94:CE158)====
    For i = 94 To 158 ' BUSQUEDA EN LA COLUMNA CE
        If wsOrigen.Cells(i, "CE").Value <> "" Then ' VERIFICA QUE LA CELDA NO ESTA VACIA
            If InStr(1, UCase(wsOrigen.Cells(i, "CE").Value), UCase(PalabraBuscada)) > 0 Then ' BUSCA LA PLABRA EN CADA TEXTO
                encontrada = True
                
                '====Determina la fila destino====
                filaDestino = 84 ' DETERMINA DONDE PEGAR
                Do While Len(wsDestino.Cells(filaDestino, "BA").Value) > 0 'verifica si la celda BA en esa fila tiene algo (longitud mayor a 0)
                    filaDestino = filaDestino + 4
                    If filaDestino > 10000 Then Exit Do
                Loop
                
                '====Copia SOLO las 4 celdas específicas====
                ' BK -> BA
                wsDestino.Cells(filaDestino, "BA").Value = wsOrigen.Cells(i, "BK").Value
                'BQ -> BG
                wsDestino.Cells(filaDestino, "BG").Value = wsOrigen.Cells(i, "BQ").Value
                ' CE -> BU
                wsDestino.Cells(filaDestino, "BU").Value = wsOrigen.Cells(i, "CE").Value
                ' CS -> CI
                wsDestino.Cells(filaDestino, "CI").Value = wsOrigen.Cells(i, "CS").Value
                ' DL -> DB
                wsDestino.Cells(filaDestino, "DB").Value = wsOrigen.Cells(i, "DL").Value
                
                '====Modifica las celdas especificadas====
                wsDestino.Cells(filaDestino, "EF").Value = MTM1
                wsDestino.Cells(filaDestino, "ER").Value = MTM2
                wsDestino.Cells(filaDestino, "FD").Value = BRR
                wsDestino.Cells(filaDestino, "GB").Value = CP
                wsDestino.Cells(filaDestino, "GZ").Value = CC
                wsDestino.Cells(filaDestino, "HX").Value = CM
            End If
        End If
    Next i
    
    '====Mensaje final====
    If encontrada Then
        MsgBox "Se encontró la palabra '" & PalabraBuscada & "' y se copió la información.", vbInformation
    Else
        MsgBox "No se encontró la palabra '" & PalabraBuscada & "'.", vbExclamation
    End If
    
    Exit Sub
    
    ErrorHandler:
        MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub 
'----------------------------------------------------------------------
' BOTONES
'----------------------------------------------------------------------
public sub T_Power()
    palabra = "T-Power"
    buscar_Inspeccion palabra , "FR-0856", "FR-0857", "FR-0858", "FR-0859", "FR-0860", "Por definir"
End Sub
public sub G_Flex()
    palabra = "G-Flex"
    buscar_Inspeccion palabra , "FR-0850", "FR-0851", "FR-0852", "FR-0853", "FR-0854", "Por definir"
End Sub
