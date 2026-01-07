'============= HOJA DE APLICACIONES ========
'=======VARIABLES GLOBALES======
public Palabra As String' SE GUARDA EL TEXTO CORRESPONDIENTE

public sub Worksheet_Activate()

    activeSheet.Shapes("Grupo I-MCC").Visible = False
    activeSheet.Shapes("Grupo L-Meter").Visible = False
    activeSheet.Shapes("Grupo O-Paralel").Visible = False
    activeSheet.Shapes("Grupo P-ATS").Visible = False
    activeSheet.Shapes("Grupo Q-Bank").Visible = False
    activeSheet.Shapes("Grupo V-Drive").Visible = False
   

end sub
Public Sub buscar_Producto(PalabraBuscada As String)
    On Error GoTo ErrorHandler

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim i As Long ' CONTADOR PARA RECCORER LAS FILAS
    Dim filaDestino As Long
    Dim encontrada As Boolean

    '====SELECCIONA LAS HOJAS DE ORIGEN Y DESTINO====
    Set wsOrigen = ThisWorkbook.Worksheets("PROYECTO")
    Set wsDestino = ThisWorkbook.Worksheets("APLICACIONES")
    encontrada = False
    
    '====BUSQUEDA DE PALABRA EN ORIGEN (CE94:CE158)====
    For i = 96 To 158 ' BUSQUEDA EN LA COLUMNA CE
        If wsOrigen.Cells(i, "CE").Value <> "" Then ' VERIFICA QUE LA CELDA NO ESTA VACIA
            If InStr(1, UCase(wsOrigen.Cells(i, "CE").Value), UCase(PalabraBuscada)) > 0 Then ' BUSCA LA PLABRA EN CADA TEXTO
                encontrada = True
                
                '====Determina la fila destino====
                filaDestino = 84 ' DETERMINA DONDE PEGAR
                Do While Len(wsDestino.Cells(filaDestino, "BA").Value) > 0 'verifica si la celda BA en esa fila tiene algo (longitud mayor a 0)
                    filaDestino = filaDestino + 4
                    If filaDestino > 10000 Then Exit Do
                Loop
                '====AGREGAMOS ESTA LÍNEA====
                wsDestino.Cells(filaDestino, "A").Value = wsOrigen.Cells(i, "BK").Value

                '====Copia SOLO las 4 celdas específicas====
                ' BK -> BA
                wsDestino.Cells(filaDestino, "BA").Value = wsOrigen.Cells(i, "BK").Value
                ' CE -> BU
                wsDestino.Cells(filaDestino, "BU").Value = wsOrigen.Cells(i, "CE").Value
                ' DL -> DB
                wsDestino.Cells(filaDestino, "DB").Value = wsOrigen.Cells(i, "DL").Value
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
    buscar_Producto palabra 
End Sub
public sub G_Flex()
    palabra = "G-Flex"
    buscar_Producto palabra
End Sub

Public Sub Cargar_IMCC()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim IMCC As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.IMCC1.Value = True Then
        T1 = "Tablero para banco de condensadores (PF)"
        'T1F = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    If ActiveSheet.IMCC2.Value = True Then
        T2 = "Tablero para filtros pasivos (HF)"
        'T2F = "???" 'no se encontro referencia del formato
    End If
    If ActiveSheet.IMCC3.Value = True Then
        T3 = "Tablero para banco de cargas resisitivos (RL)"
        'T3F = "FR-0811"'cargado
    End If
    If ActiveSheet.IMCC4.Value = True Then
        T4 = "Compensadores dinámicos (CD)"
        'T4F = "???" 'no se encontro referencia del formato
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    IMCC = ""
    If T1 <> "" Then IMCC = IMCC & T1 & "/"
    If T2 <> "" Then IMCC = IMCC & T2 & "/"
    If T3 <> "" Then IMCC = IMCC & T3 & "/"
    If T4 <> "" Then IMCC = IMCC & T4 & "/"
    If T5 <> "" Then IMCC = IMCC & T5
    
    ' Limpiar saltos de línea finales
    IMCC = Trim(IMCC)
    If Right(IMCC, 1) = Chr(10) Then
        IMCC = Left(IMCC, Len(IMCC) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If IMCC = "" Then
        MsgBox "No se ha seleccionado ningún Formato de la Aplicación.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de la Aplicación:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA (columna 53)
    ultimaFila = ws.Cells(ws.Rows.Count, "BA").End(xlUp).Row
    encontrado = False
    
    For filaOT = 1 To ultimaFila
        If Trim(CStr(ws.Cells(filaOT, "BA").Value)) = Trim(OT_Buscada) Then
            encontrado = True
            Exit For
        End If
    Next filaOT
    
    ' Si se encontró la OT, escribir en la columna EF
    If encontrado Then
        ' Escribir en la columna EF (columna 136) - CORREGIDO: Ahora escribe Ensayos_Personalizado
        ws.Cells(filaOT, "EF").Value = IMCC
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & IMCC, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_LMETER()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim LMETER As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    Set ws = ActiveSheet
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.LMeter1.Value = True Then
        T1 = "Tableros para medidores de energía ( D - S.D - P.M )"
        'T1F = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    ' Construir el texto de ensayos (eliminar líneas vacías)
    LMETER = ""
    If T1 <> "" Then LMETER = LMETER & T1
    ' Limpiar saltos de línea finales
    LMETER = Trim(LMETER)
    If Right(LMETER, 1) = Chr(10) Then
        LMETER = Left(LMETER, Len(LMETER) - 1)
    End If
    ' Verificar que hay ensayos seleccionados
    If LMETER = "" Then
        MsgBox "No se ha seleccionado ningún Formato de la Aplicación.", vbExclamation, "Aviso"
        Exit Sub
    End If
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de la Aplicación:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If  
    ' Buscar la OT en la columna BA (columna 53)
    ultimaFila = ws.Cells(ws.Rows.Count, "BA").End(xlUp).Row
    encontrado = False
    For filaOT = 1 To ultimaFila
        If Trim(CStr(ws.Cells(filaOT, "BA").Value)) = Trim(OT_Buscada) Then
            encontrado = True
            Exit For
        End If
    Next filaOT
    ' Si se encontró la OT, escribir en la columna EF
    If encontrado Then
        ' Escribir en la columna EF (columna 136) - CORREGIDO: Ahora escribe Ensayos_Personalizado
        ws.Cells(filaOT, "EF").Value = LMETER
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & LMETER, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
End Sub
Public Sub Cargar_OPARALEL()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim OPARALEL As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    Set ws = ActiveSheet
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.OParalel1.Value = True Then
        T1 = "Tableros de sincronismo"
        'T1F = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    ' Construir el texto de ensayos (eliminar líneas vacías)
    OPARALEL = ""
    If T1 <> "" Then OPARALEL = OPARALEL & T1
    ' Limpiar saltos de línea finales
    OPARALEL = Trim(OPARALEL)
    If Right(OPARALEL, 1) = Chr(10) Then
        OPARALEL = Left(OPARALEL, Len(OPARALEL) - 1)
    End If
    ' Verificar que hay ensayos seleccionados
    If OPARALEL = "" Then
        MsgBox "No se ha seleccionado ningún Formato de la Aplicación.", vbExclamation, "Aviso"
        Exit Sub
    End If
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de la Aplicación:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If  
    ' Buscar la OT en la columna BA (columna 53)
    ultimaFila = ws.Cells(ws.Rows.Count, "BA").End(xlUp).Row
    encontrado = False
    For filaOT = 1 To ultimaFila
        If Trim(CStr(ws.Cells(filaOT, "BA").Value)) = Trim(OT_Buscada) Then
            encontrado = True
            Exit For
        End If
    Next filaOT
    ' Si se encontró la OT, escribir en la columna EF
    If encontrado Then
        ' Escribir en la columna EF (columna 136) - CORREGIDO: Ahora escribe Ensayos_Personalizado
        ws.Cells(filaOT, "EF").Value = OPARALEL
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & OPARALEL, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
End Sub
Public Sub Cargar_IMCC()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim IMCC As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.IMCC1.Value = True Then
        T1 = "Tablero para banco de condensadores (PF)"
        'T1F = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    If ActiveSheet.IMCC2.Value = True Then
        T2 = "Tablero para filtros pasivos (HF)"
        'T2F = "???" 'no se encontro referencia del formato
    End If
    If ActiveSheet.IMCC3.Value = True Then
        T3 = "Tablero para banco de cargas resisitivos (RL)"
        'T3F = "FR-0811"'cargado
    End If
    If ActiveSheet.IMCC4.Value = True Then
        T4 = "Compensadores dinámicos (CD)"
        'T4F = "???" 'no se encontro referencia del formato
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    IMCC = ""
    If T1 <> "" Then IMCC = IMCC & T1 & "/"
    If T2 <> "" Then IMCC = IMCC & T2 & "/"
    If T3 <> "" Then IMCC = IMCC & T3 & "/"
    If T4 <> "" Then IMCC = IMCC & T4 & "/"
    If T5 <> "" Then IMCC = IMCC & T5
    
    ' Limpiar saltos de línea finales
    IMCC = Trim(IMCC)
    If Right(IMCC, 1) = Chr(10) Then
        IMCC = Left(IMCC, Len(IMCC) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If IMCC = "" Then
        MsgBox "No se ha seleccionado ningún Formato de la Aplicación.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de la Aplicación:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA (columna 53)
    ultimaFila = ws.Cells(ws.Rows.Count, "BA").End(xlUp).Row
    encontrado = False
    
    For filaOT = 1 To ultimaFila
        If Trim(CStr(ws.Cells(filaOT, "BA").Value)) = Trim(OT_Buscada) Then
            encontrado = True
            Exit For
        End If
    Next filaOT
    
    ' Si se encontró la OT, escribir en la columna EF
    If encontrado Then
        ' Escribir en la columna EF (columna 136) - CORREGIDO: Ahora escribe Ensayos_Personalizado
        ws.Cells(filaOT, "EF").Value = IMCC
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & IMCC, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub