'==================== HOJA DE EQUIPOS =======================
'=======VARIABLES GLOBALES======
public Palabra As String' SE GUARDA EL TEXTO CORRESPONDIENTE


Public sub limpiar_tabla_EQUIPOS()
    'LIMPIAR VARIABLES
    AP1 = ""
    AP2 = ""
    AP3 = ""
    AP4 = ""
    AP5 = ""
    AP6 = ""
    AP7 = ""
    AP8 = ""

    ThisWorkbook.Sheets("EQUIPOS").range("BA84:EF140").Value = ""
end sub
     
Public sub Worksheet_Activate()
    activeSheet.Shapes("Sistema de proteccion").Visible = False
    activeSheet.Shapes("Conversion y distribucion de energia").Visible = False
    activeSheet.Shapes("Dispositivos de supervision y control").Visible = False
    activeSheet.Shapes("Instrumentacion y medicion").Visible = False
    activeSheet.Shapes("Servicios Auxiliares").Visible = False
end sub
Public Sub buscar_Producto(PalabraBuscada As String)
    On Error GoTo ErrorHandler

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim i As Long ' CONTADOR PARA RECCORER LAS FILAS
    Dim filaDestino As Long
    Dim encontrada As Boolean

    '====SELECCIONA LAS HOJAS DE ORIGEN Y DESTINO====
    Set wsOrigen = ThisWorkbook.Worksheets("PROYECTO")
    Set wsDestino = ThisWorkbook.Worksheets("EQUIPOS")
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
'----------------------------------------------------------------------

Public Sub Cargar_SP()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim SP As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Inicializar variables
    T1 = ""
    T2 = ""
    T3 = ""
    T4 = "" 
    T5 = ""
    EQ1 = ""
    EQ2 = ""
    EQ3 = ""
    EQ4 = ""
    
    
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("SP1").Object.Value = True Then
            T1 = "Prueba a unidad de disparo Micrologic S.E"
            'EQ1 = "FR-0829"
        End If
        If ws.OLEObjects("SP2").Object.Value = True Then
            T2 = "Dispositivo de protección contra sobretenciones (DPS)"
            'EQ2 = "FR-0841"
        End If
        If ws.OLEObjects("SP3").Object.Value = True Then
            T3 = "Interruptores diferenciales"
            'EQ3 = "FR-0811"
        End If
        On Error GoTo 0
    End If
    
    ' Construir el texto de SP
    SP = ""
    If T1 <> "" Then SP = SP & T1 & "/"
    If T2 <> "" Then SP = SP & T2 & "/"
    If T3 <> "" Then SP = SP & T3 
    
    ' Eliminar el último "/" si existe
    If Right(SP, 1) = "/" Then
        SP = Left(SP, Len(SP) - 1)
    End If
    
    ' Verificar que hay formatos seleccionados
    If SP = "" Then 
        MsgBox "No se ha seleccionado ningún Formato para los equipos.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de los equipos:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA
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
        ws.Cells(filaOT, "EF").Value = SP
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de equipos:" & vbCrLf & SP, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & ". Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_CD()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim CD As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Inicializar variables
        T1 = ""
        T2 = ""
        T3 = ""
        T4 = "" 
        T5 = ""
        EQ1 = ""
        EQ2 = ""
        EQ3 = ""
        EQ4 = ""
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("CD1").Object.Value = True Then
            T1 = "Transformador de potencia"
            'EQ1 = "FR-0842"
        End If
        If ws.OLEObjects("CD2").Object.Value = True Then
            T2 = "Transformador de control"
            'EQ2 = "FR-0844"
        End If
        If ws.OLEObjects("CD3").Object.Value = True Then
            T3 = "Sistema de alimentación initerrumpida (UPS)"
            'EQ3 = "FR-0843"
        End If
        On Error GoTo 0
    End If
    
    ' Construir el texto de CD
    CD = ""
    If T1 <> "" Then CD = CD & T1 & "/"
    If T2 <> "" Then CD = CD & T2 & "/"
    If T3 <> "" Then CD = CD & T3 
    
    ' Eliminar el último "/" si existe
    If Right(CD, 1) = "/" Then
        CD = Left(CD, Len(CD) - 1)
    End If
    
    ' Verificar que hay formatos seleccionados
    If CD = "" Then 
        MsgBox "No se ha seleccionado ningún Formato para los equipos.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de los equipos:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA
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
        ws.Cells(filaOT, "EF").Value = CD
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de equipos:" & vbCrLf & CD, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & ". Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_IM()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim IM As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Inicializar variables
        T1 = ""
        T2 = ""
        T3 = ""
        T4 = "" 
        T5 = ""
        EQ1 = ""
        EQ2 = ""
        EQ3 = ""
        EQ4 = ""
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("CD1").Object.Value = True Then
            T1 = "Medidores (directo-semidirecto-indirecto)"
            'EQ1 = "FR-0813"
        End If
        If ws.OLEObjects("CD2").Object.Value = True Then
            T2 = "Conjunto de CT'S"
            'EQ2 = "FR-0848"
        End If
        If ws.OLEObjects("CD3").Object.Value = True Then
            T3 = "Conjunto de PT'S"
            'EQ3 = "FR-0848"
        End If
        On Error GoTo 0
    End If
    
    ' Construir el texto de IM
    IM = ""
    If T1 <> "" Then IM = IM & T1 & "/"
    If T2 <> "" Then IM = IM & T2 & "/"
    If T3 <> "" Then IM = IM & T3 
    
    ' Eliminar el último "/" si existe
    If Right(IM, 1) = "/" Then
        IM = Left(IM, Len(IM) - 1)
    End If
    
    ' Verificar que hay formatos seleccionados
    If IM = "" Then 
        MsgBox "No se ha seleccionado ningún Formato para los equipos.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de los equipos:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA
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
        ws.Cells(filaOT, "EF").Value = IM
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de equipos:" & vbCrLf & IM, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & ". Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_SA()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim SA As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Inicializar variables
        T1 = ""
        T2 = ""
        T3 = ""
        T4 = "" 
        T5 = ""
        EQ1 = ""
        EQ2 = ""
        EQ3 = ""
        EQ4 = ""
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("CD1").Object.Value = True Then
            T1 = "Ventilación"
            'EQ1 = "FR-0849"
        End If
        If ws.OLEObjects("CD2").Object.Value = True Then
            T2 = "Iluminación"
            'EQ2 = "???"
        End If
        If ws.OLEObjects("CD3").Object.Value = True Then
            T3 = "Calefacción"
            'EQ3 = "???"
        End If
        If ws.OLEObjects("CD4").Object.Value = True Then
            T4 = "Toma de servicios"
            'EQ4 = "???"
        End If
        On Error GoTo 0
    End If
    
    ' Construir el texto de SA
    SA = ""
    If T1 <> "" Then SA = SA & T1 & "/"
    If T2 <> "" Then SA = SA & T2 & "/"
    If T3 <> "" Then SA = SA & T3 
    
    ' Eliminar el último "/" si existe
    If Right(SA, 1) = "/" Then
        SA = Left(SA, Len(SA) - 1)
    End If
    
    ' Verificar que hay formatos seleccionados
    If SA = "" Then 
        MsgBox "No se ha seleccionado ningún Formato para los equipos.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los Formatos de los equipos:", _
                          "Buscar OT")
    
    ' Validar que se ingresó una OT
    If OT_Buscada = "" Then
        MsgBox "Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA
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
        ws.Cells(filaOT, "EF").Value = SA
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de equipos:" & vbCrLf & SA, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & ". Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub