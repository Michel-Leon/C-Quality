'==========HOJA DE ENSAYOS ===============
'=======VARIABLES GLOBALES======
public Palabra As String' SE GUARDA EL TEXTO CORRESPONDIENTE
Public PaqueteBasico as String, PaqueteEstandar as String, PaquetePro as String, PaqueteEnterprise as String, PaquetePersonalizado as String


public sub Worksheet_Activate()

    with Me.OLEObjects("Ensayo1").Object
        .Clear 
        .AddItem "Efectividad del circuito de protección."
        .AddItem "Espesor de pintura."
        .AddItem "Resistencia de aislamiento."
        .AddItem "Operación y funcionamiento mecánico."
        .AddItem "Inyección de corriente primaria por inducción (Circuito de medida.)"
        .AddItem "Tensión soportada a frecuencia industrial."
        .AddItem "Adherencia de pintura."
        .AddItem "Grado de protección IP (Hasta 4X)."
        .AddItem "Imborabilidad del Rotulado."
        .AddItem "Permanecía del Rotulado."
        .AddItem "Pruebas de unidad de disparo Micrologic SE."
        .AddItem "Tiempos de Disparo y Cierre."
        .AddItem "Discrepancia de polos."
        .AddItem "Test de verificación y operación Interruptor NSX."
        .AddItem "Conversión y distribución de energía."
        .AddItem "Dispositivos de supervisión y control."
        .AddItem "Instrumentación y medición."
        .AddItem "Servicios Auxiliares."
    end with
    with Me.OLEObjects("Ensayo2").Object
        .Clear 
        .AddItem "Efectividad del circuito de protección."
        .AddItem "Espesor de pintura."
        .AddItem "Resistencia de aislamiento."
        .AddItem "Operación y funcionamiento mecánico."
        .AddItem "Inyección de corriente primaria por inducción (Circuito de medida.)"
        .AddItem "Tensión soportada a frecuencia industrial."
        .AddItem "Adherencia de pintura."
        .AddItem "Grado de protección IP (Hasta 4X)."
        .AddItem "Imborabilidad del Rotulado."
        .AddItem "Permanecía del Rotulado."
        .AddItem "Pruebas de unidad de disparo Micrologic SE."
        .AddItem "Tiempos de Disparo y Cierre."
        .AddItem "Discrepancia de polos."
        .AddItem "Test de verificación y operación Interruptor NSX."
        .AddItem "Conversión y distribución de energía."
        .AddItem "Dispositivos de supervisión y control."
        .AddItem "Instrumentación y medición."
        .AddItem "Servicios Auxiliares."
    end with
    with Me.OLEObjects("Ensayo3").Object
        .Clear 
        .AddItem "Efectividad del circuito de protección."
        .AddItem "Espesor de pintura."
        .AddItem "Resistencia de aislamiento."
        .AddItem "Operación y funcionamiento mecánico."
        .AddItem "Inyección de corriente primaria por inducción (Circuito de medida.)"
        .AddItem "Tensión soportada a frecuencia industrial."
        .AddItem "Adherencia de pintura."
        .AddItem "Grado de protección IP (Hasta 4X)."
        .AddItem "Imborabilidad del Rotulado."
        .AddItem "Permanecía del Rotulado."
        .AddItem "Pruebas de unidad de disparo Micrologic SE."
        .AddItem "Tiempos de Disparo y Cierre."
        .AddItem "Discrepancia de polos."
        .AddItem "Test de verificación y operación Interruptor NSX."
        .AddItem "Conversión y distribución de energía."
        .AddItem "Dispositivos de supervisión y control."
        .AddItem "Instrumentación y medición."
        .AddItem "Servicios Auxiliares."
    end with
    with Me.OLEObjects("Ensayo4").Object
        .Clear 
        .AddItem "Efectividad del circuito de protección."
        .AddItem "Espesor de pintura."
        .AddItem "Resistencia de aislamiento."
        .AddItem "Operación y funcionamiento mecánico."
        .AddItem "Inyección de corriente primaria por inducción (Circuito de medida.)"
        .AddItem "Tensión soportada a frecuencia industrial."
        .AddItem "Adherencia de pintura."
        .AddItem "Grado de protección IP (Hasta 4X)."
        .AddItem "Imborabilidad del Rotulado."
        .AddItem "Permanecía del Rotulado."
        .AddItem "Pruebas de unidad de disparo Micrologic SE."
        .AddItem "Tiempos de Disparo y Cierre."
        .AddItem "Discrepancia de polos."
        .AddItem "Test de verificación y operación Interruptor NSX."
        .AddItem "Conversión y distribución de energía."
        .AddItem "Dispositivos de supervisión y control."
        .AddItem "Instrumentación y medición."
        .AddItem "Servicios Auxiliares."
    end with
    with Me.OLEObjects("Ensayo5").Object
        .Clear 
        .AddItem "Efectividad del circuito de protección."
        .AddItem "Espesor de pintura."
        .AddItem "Resistencia de aislamiento."
        .AddItem "Operación y funcionamiento mecánico."
        .AddItem "Inyección de corriente primaria por inducción (Circuito de medida.)"
        .AddItem "Tensión soportada a frecuencia industrial."
        .AddItem "Adherencia de pintura."
        .AddItem "Grado de protección IP (Hasta 4X)."
        .AddItem "Imborabilidad del Rotulado."
        .AddItem "Permanecía del Rotulado."
        .AddItem "Pruebas de unidad de disparo Micrologic SE."
        .AddItem "Tiempos de Disparo y Cierre."
        .AddItem "Discrepancia de polos."
        .AddItem "Test de verificación y operación Interruptor NSX."
        .AddItem "Conversión y distribución de energía."
        .AddItem "Dispositivos de supervisión y control."
        .AddItem "Instrumentación y medición."
        .AddItem "Servicios Auxiliares."
    end with

    'Activar los checkboxes --------------------------------------------
        ActiveSheet.B1.Value = True
        ActiveSheet.B2.Value = True
        ActiveSheet.B3.Value = True
        ActiveSheet.B4.Value = True
        ActiveSheet.B5.Value = True
        ActiveSheet.ES1.Value = True
        ActiveSheet.ES2.Value = True
        ActiveSheet.ES3.Value = True
        ActiveSheet.ES4.Value = True
        ActiveSheet.ES5.Value = True
        ActiveSheet.PRO1.Value = True
        ActiveSheet.PRO2.Value = True
        ActiveSheet.PRO3.Value = True
        ActiveSheet.PRO4.Value = True
        ActiveSheet.EN1.Value = True
        ActiveSheet.EN2.Value = True
        ActiveSheet.EN3.Value = True
        ActiveSheet.EN4.Value = True
    '----------------------------------------------------------------------
end sub    
'----------------------------------------------------------------------
'                   PROCEDIMIENTOS
'----------------------------------------------------------------------
Public Sub buscar_ensayos(PalabraBuscada As String)
    On Error GoTo ErrorHandler

    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim i As Long ' CONTADOR PARA RECCORER LAS FILAS
    Dim filaDestino As Long
    Dim encontrada As Boolean

    '====SELECCIONA LAS HOJAS DE ORIGEN Y DESTINO====
    Set wsOrigen = ThisWorkbook.Worksheets("PROYECTO")
    Set wsDestino = ThisWorkbook.Worksheets("ENSAYOS R")
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
    buscar_ensayos palabra 
End Sub
public sub G_Flex()
    palabra = "G-Flex"
    buscar_ensayos palabra
End Sub
'----------------------------------------------------------------------
' CARGAR ENSAYOS
'----------------------------------------------------------------------
Public Sub Cargar_Basico()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim ENSAYOS_Basico As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.B1.Value = True Then
        T1 = "Efectividad del circuito de protección."
        T1F = "FR-0804" 'cargado
    End If
    If ActiveSheet.B2.Value = True Then
        T2 = "Espesor de pintura."
        T2F = "FR-0815" 'no se encontro referencia del formato
    End If
    If ActiveSheet.B3.Value = True Then
        T3 = "Resistencia de aislamiento."
        T3F = "FR-0811"'cargado
    End If
    If ActiveSheet.B4.Value = True Then
        T4 = "Operación y funcionamiento mecánico."
        T4F = "FR-0890" 'no se encontro referencia del formato
    End If
    If ActiveSheet.B5.Value = True Then
        T5 = "Inyección de corriente primaria por inducción (Circuito de medida.)"
        T5F = "???" 'no se encontro referencia del formato
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    ENSAYOS_Basico = ""
    If T1 <> "" Then ENSAYOS_Basico = ENSAYOS_Basico & T1 & "/"
    If T2 <> "" Then ENSAYOS_Basico = ENSAYOS_Basico & T2 & "/"
    If T3 <> "" Then ENSAYOS_Basico = ENSAYOS_Basico & T3 & "/"
    If T4 <> "" Then ENSAYOS_Basico = ENSAYOS_Basico & T4 & "/"
    If T5 <> "" Then ENSAYOS_Basico = ENSAYOS_Basico & T5
    
    ' Limpiar saltos de línea finales
    ENSAYOS_Basico = Trim(ENSAYOS_Basico)
    If Right(ENSAYOS_Basico, 1) = Chr(10) Then
        ENSAYOS_Basico = Left(ENSAYOS_Basico, Len(ENSAYOS_Basico) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If ENSAYOS_Basico = "" Then
        MsgBox "No se ha seleccionado ningún ensayo básico.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los ensayos básicos:", _
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
        ws.Cells(filaOT, "EF").Value = Ensayos_Basico
        
        MsgBox "Ensayos  cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Ensayos:" & vbCrLf & Ensayos_Basico, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " en la columna BA.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub

public Sub Cargar_estandar()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim Ensayos_Estandar As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.ES1.Value = True Then
        T1 = "Tensión soportada a frecuencia industrial."
        T1F = "FR-0803" 'cargado
    End If
    If ActiveSheet.ES2.Value = True Then
        T2 = "Adherencia de pintura."
        T2F = "FR-0815" 'Cargado
    End If
    If ActiveSheet.ES3.Value = True Then
        T3 = "Grado de protección IP (Hasta 4X)."
        T3F = "FR-0811"'cargado
    End If
    If ActiveSheet.ES4.Value = True Then
        T4 = "Imborabilidad del Rotulado."
        T4F = "FR-0890" 'nCargado
    End If
    If ActiveSheet.ES5.Value = True Then
        T5 = "Permanecía del Rotulado."
        T5F = "???" 'no se encontro referencia del formato
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    Ensayos_Estandar = ""
    If T1 <> "" Then Ensayos_Estandar = Ensayos_Estandar & T1 & "/"
    If T2 <> "" Then Ensayos_Estandar = Ensayos_Estandar & T2 & "/"
    If T3 <> "" Then Ensayos_Estandar = Ensayos_Estandar & T3 & "/"
    If T4 <> "" Then Ensayos_Estandar = Ensayos_Estandar & T4 & "/"
    If T5 <> "" Then Ensayos_Estandar = Ensayos_Estandar & T5
    
    ' Limpiar saltos de línea finales
    Ensayos_Estandar = Trim(Ensayos_Estandar)
    If Right(Ensayos_Estandar, 1) = Chr(10) Then
        Ensayos_Estandar = Left(Ensayos_Estandar, Len(Ensayos_Estandar) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If Ensayos_Estandar = "" Then
        MsgBox "No se ha seleccionado ningún ensayo estandar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los ensayos estandar:", _
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
        ws.Cells(filaOT, "EF").Value = Ensayos_Estandar
        
        MsgBox "Ensayos  cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Ensayos:" & vbCrLf & Ensayos_Estandar, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " en la columna BA.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub

public Sub Cargar_PRO()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim Ensayos_pro As String
    Dim ws As Worksheet
    Dim OT_Buscada As String
    Dim filaOT As Long
    Dim ultimaFila As Long
    Dim encontrado As Boolean
    
    Set ws = ActiveSheet
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.ES1.Value = True Then
        T1 = "Pruebas a unidad de disparo Micrologic SE."
        T1F = "FR-0829" 'cargado
    End If
    If ActiveSheet.ES2.Value = True Then
        T2 = "Tiempos de disparo y cierre."
        T2F = "???" 'NO SE ENCONTRO REFERENCIA
    End If
    If ActiveSheet.ES3.Value = True Then
        T3 = "Discrepancia de polos."
        T3F = "???"'NO SE ENCONTRO REFERENCIA
    End If
    If ActiveSheet.ES4.Value = True Then
        T4 = "Test verificación y operacion interruptor NSX."
        T4F = "???" 'NO SE ENCONTRO REFERENCIA
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    Ensayos_pro = ""
    If T1 <> "" Then Ensayos_pro = Ensayos_pro & T1 & "/"
    If T2 <> "" Then Ensayos_pro = Ensayos_pro & T2 & "/"
    If T3 <> "" Then Ensayos_pro = Ensayos_pro & T3 & "/"
    If T4 <> "" Then Ensayos_pro = Ensayos_pro & T4 & "/"
    If T5 <> "" Then Ensayos_pro = Ensayos_pro & T5
    
    ' Limpiar saltos de línea finales
    Ensayos_pro = Trim(Ensayos_pro)
    If Right(Ensayos_pro, 1) = Chr(10) Then
        Ensayos_pro = Left(Ensayos_pro, Len(Ensayos_pro) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If Ensayos_pro = "" Then
        MsgBox "No se ha seleccionado ningún ensayo estandar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los ensayos estandar:", _
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
        ws.Cells(filaOT, "EF").Value = Ensayos_pro
        
        MsgBox "Ensayos  cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Ensayos:" & vbCrLf & Ensayos_pro, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " en la columna BA.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub

Public Sub Cargar_Enterprise()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim Ensayos_Enterprise As String
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
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.EN1.Value = True Then
        T1 = "Conversión y distribución de energia"
        T1F = "???" 'no se encontro referencia
    End If
    If ActiveSheet.EN2.Value = True Then
        T2 = "Disposicion de supervisión y control"
        T2F = "???" 'NO SE ENCONTRO REFERENCIA
    End If
    If ActiveSheet.EN3.Value = True Then
        T3 = "Instrumentación y medición"
        T3F = "???"'NO SE ENCONTRO REFERENCIA
    End If
    If ActiveSheet.EN4.Value = True Then
        T4 = "Servicios Auxiliares"
        T4F = "???" 'NO SE ENCONTRO REFERENCIA
    End If
    
    ' Construir el texto de ensayos
    Ensayos_Enterprise = ""
    If T1 <> "" Then Ensayos_Enterprise = Ensayos_Enterprise & T1 & "/"
    If T2 <> "" Then Ensayos_Enterprise = Ensayos_Enterprise & T2 & "/"
    If T3 <> "" Then Ensayos_Enterprise = Ensayos_Enterprise & T3 & "/"
    If T4 <> "" Then Ensayos_Enterprise = Ensayos_Enterprise & T4 & "/"
    If T5 <> "" Then Ensayos_Enterprise = Ensayos_Enterprise & T5 & "/"
    
    ' Eliminar el último "/" si existe
    If Right(Ensayos_Enterprise, 1) = "/" Then
        Ensayos_Enterprise = Left(Ensayos_Enterprise, Len(Ensayos_Enterprise) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If Ensayos_Enterprise = "" Then
        MsgBox "No se ha seleccionado ningún ensayo Enterprise.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los ensayos Enterprise:", _
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
        ' Escribir en la columna EF (columna 136) - CORREGIDO: Era ENSAYOS_Basico, ahora es Ensayos_Enterprise
        ws.Cells(filaOT, "EF").Value = Ensayos_Enterprise
        
        MsgBox "Ensayos  cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Ensayos:" & vbCrLf & Ensayos_Enterprise, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " en la columna BA.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub

Public Sub Cargar_Personalizado()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim Ensayos_Personalizado As String
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
    
    ' Verificar que los combobox tienen valor - ENSAYO 1
    Select Case Me.OLEObjects("Ensayo1").Object.Value
        Case "Efectividad del circuito de protección."
            T1 = "Efectividad del circuito de protección"
            T1F = "FR-0804"
        Case "Espesor de pintura."
            T1 = "Espesor de pintura"
            T1F = "FR-0815"
        Case "Resistencia de aislamiento."
            T1 = "Resistencia de aislamiento"
            T1F = "FR-0811"
        Case "Operación y funcionamiento mecánico."
            T1 = "Operación y funcionamiento mecánico"
            T1F = "FR-0890"
        Case "Instrumentacion y medicion."
            T1 = "Instrumentacion y medicion"
            T1F = "???"
        Case "Conversión y distribución de energia."
            T1 = "Conversión y distribución de energia"
            T1F = "???"
        Case "Disposicion de supervisoon y control."
            T1 = "Disposicion de supervisoon y control"
            T1F = "???"
        Case "Inyección de corriente primaria por inducción (Circuito de medida.)"
            T1 = "Inyección de corriente primaria por inducción (Circuito de medida)"
            T1F = "???"
        Case "Tensión soportada a frecuencia industrial."
            T1 = "Tensión soportada a frecuencia industrial"
            T1F = "FR-0803"
        Case "Adherencia de pintura."
            T1 = "Adherencia de pintura"
            T1F = "???"
        Case "Grado de protección IP (Hasta 4X)."
            T1 = "Grado de protección IP (Hasta 4X)"
            T1F = "???"
        Case "Imborabilidad del Rotulado."
            T1 = "Imborabilidad del Rotulado"
            T1F = "???"
        Case "Permanecía del Rotulado."
            T1 = "Permanecía del Rotulado"
            T1F = "???"
        Case "Pruebas a unidad de disparo Micrologic SE."
            T1 = "Pruebas a unidad de disparo Micrologic SE"
            T1F = "FR-0829"
        Case "Tiempos de disparo y cierre."
            T1 = "Tiempos de disparo y cierre"
            T1F = "FR-0829"
        Case "Discrepancia de polos."
            T1 = "Discrepancia de polos"
            T1F = "???"
        Case "Test de verificación y operación Interruptor NSX."
            T1 = "Test de verificación y operación Interruptor NSX"
            T1F = "???"
        Case "Servicios Auxiliares."
            T1 = "Servicios Auxiliares"
            T1F = "???"
        Case "Dispositivos de supervisión y control"
            T1 = "Dispositivos de supervisión y control"
            T1F = "???"
    End Select
    
    ' Verificar que los combobox tienen valor - ENSAYO 2
    Select Case Me.OLEObjects("Ensayo2").Object.Value
        Case "Efectividad del circuito de protección."
            T2 = "Efectividad del circuito de protección"
            T2F = "FR-0804"
        Case "Espesor de pintura."
            T2 = "Espesor de pintura"
            T2F = "FR-0815"
        Case "Resistencia de aislamiento."
            T2 = "Resistencia de aislamiento"
            T2F = "FR-0811"
        Case "Operación y funcionamiento mecánico."
            T2 = "Operación y funcionamiento mecánico"
            T2F = "FR-0890"
        Case "Instrumentacion y medicion."
            T2 = "Instrumentacion y medicion"
            T2F = "???"
        Case "Conversión y distribución de energia."
            T2 = "Conversión y distribución de energia"
            T2F = "???"
        Case "Disposicion de supervisoon y control."
            T2 = "Disposicion de supervisoon y control"
            T2F = "???"
        Case "Inyección de corriente primaria por inducción (Circuito de medida.)"
            T2 = "Inyección de corriente primaria por inducción (Circuito de medida)"
            T2F = "???"
        Case "Tensión soportada a frecuencia industrial."
            T2 = "Tensión soportada a frecuencia industrial"
            T2F = "FR-0803"
        Case "Adherencia de pintura."
            T2 = "Adherencia de pintura"
            T2F = "???"
        Case "Grado de protección IP (Hasta 4X)."
            T2 = "Grado de protección IP (Hasta 4X)"
            T2F = "???"
        Case "Imborabilidad del Rotulado."
            T2 = "Imborabilidad del Rotulado"
            T2F = "???"
        Case "Permanecía del Rotulado."
            T2 = "Permanecía del Rotulado"
            T2F = "???"
        Case "Pruebas a unidad de disparo Micrologic SE."
            T2 = "Pruebas a unidad de disparo Micrologic SE"
            T2F = "FR-0829"
        Case "Tiempos de disparo y cierre."
            T2 = "Tiempos de disparo y cierre"
            T2F = "FR-0829"
        Case "Discrepancia de polos."
            T2 = "Discrepancia de polos"
            T2F = "???"
        Case "Test de verificación y operación Interruptor NSX."
            T2 = "Test de verificación y operación Interruptor NSX"
            T2F = "???"
        Case "Servicios Auxiliares."
            T2 = "Servicios Auxiliares"
            T2F = "???"
        Case "Dispositivos de supervisión y control"
            T2 = "Dispositivos de supervisión y control"
            T2F = "???"
    End Select
    
    ' Verificar que los combobox tienen valor - ENSAYO 3
    Select Case Me.OLEObjects("Ensayo3").Object.Value
        Case "Efectividad del circuito de protección."
            T3 = "Efectividad del circuito de protección"
            T3F = "FR-0804"
        Case "Espesor de pintura."
            T3 = "Espesor de pintura"
            T3F = "FR-0815"
        Case "Resistencia de aislamiento."
            T3 = "Resistencia de aislamiento"
            T3F = "FR-0811"
        Case "Operación y funcionamiento mecánico."
            T3 = "Operación y funcionamiento mecánico"
            T3F = "FR-0890"
        Case "Instrumentacion y medicion."
            T3 = "Instrumentacion y medicion"
            T3F = "???"
        Case "Conversión y distribución de energia."
            T3 = "Conversión y distribución de energia"
            T3F = "???"
        Case "Disposicion de supervisoon y control."
            T3 = "Disposicion de supervisoon y control"
            T3F = "???"
        Case "Inyección de corriente primaria por inducción (Circuito de medida.)"
            T3 = "Inyección de corriente primaria por inducción (Circuito de medida)"
            T3F = "???"
        Case "Tensión soportada a frecuencia industrial."
            T3 = "Tensión soportada a frecuencia industrial"
            T3F = "FR-0803"
        Case "Adherencia de pintura."
            T3 = "Adherencia de pintura"
            T3F = "???"
        Case "Grado de protección IP (Hasta 4X)."
            T3 = "Grado de protección IP (Hasta 4X)"
            T3F = "???"
        Case "Imborabilidad del Rotulado."
            T3 = "Imborabilidad del Rotulado"
            T3F = "???"
        Case "Permanecía del Rotulado."
            T3 = "Permanecía del Rotulado"
            T3F = "???"
        Case "Pruebas a unidad de disparo Micrologic SE."
            T3 = "Pruebas a unidad de disparo Micrologic SE"
            T3F = "FR-0829"
        Case "Tiempos de disparo y cierre."
            T3 = "Tiempos de disparo y cierre"
            T3F = "FR-0829"
        Case "Discrepancia de polos."
            T3 = "Discrepancia de polos"
            T3F = "???"
        Case "Test de verificación y operación Interruptor NSX."
            T3 = "Test de verificación y operación Interruptor NSX"
            T3F = "???"
        Case "Servicios Auxiliares."
            T3 = "Servicios Auxiliares"
            T3F = "???"
        Case "Dispositivos de supervisión y control"
            T3 = "Dispositivos de supervisión y control"
            T3F = "???"
    End Select
    
    ' Verificar que los combobox tienen valor - ENSAYO 4
    Select Case Me.OLEObjects("Ensayo4").Object.Value
        Case "Efectividad del circuito de protección."
            T4 = "Efectividad del circuito de protección"
            T4F = "FR-0804"
        Case "Espesor de pintura."
            T4 = "Espesor de pintura"
            T4F = "FR-0815"
        Case "Resistencia de aislamiento."
            T4 = "Resistencia de aislamiento"
            T4F = "FR-0811"
        Case "Operación y funcionamiento mecánico."
            T4 = "Operación y funcionamiento mecánico"
            T4F = "FR-0890"
        Case "Instrumentacion y medicion."
            T4 = "Instrumentacion y medicion"
            T4F = "???"
        Case "Conversión y distribución de energia."
            T4 = "Conversión y distribución de energia"
            T4F = "???"
        Case "Disposicion de supervisoon y control."
            T4 = "Disposicion de supervisoon y control"
            T4F = "???"
        Case "Inyección de corriente primaria por inducción (Circuito de medida.)"
            T4 = "Inyección de corriente primaria por inducción (Circuito de medida)"
            T4F = "???"
        Case "Tensión soportada a frecuencia industrial."
            T4 = "Tensión soportada a frecuencia industrial"
            T4F = "FR-0803"
        Case "Adherencia de pintura."
            T4 = "Adherencia de pintura"
            T4F = "???"
        Case "Grado de protección IP (Hasta 4X)."
            T4 = "Grado de protección IP (Hasta 4X)"
            T4F = "???"
        Case "Imborabilidad del Rotulado."
            T4 = "Imborabilidad del Rotulado"
            T4F = "???"
        Case "Permanecía del Rotulado."
            T4 = "Permanecía del Rotulado"
            T4F = "???"
        Case "Pruebas a unidad de disparo Micrologic SE."
            T4 = "Pruebas a unidad de disparo Micrologic SE"
            T4F = "FR-0829"
        Case "Tiempos de disparo y cierre."
            T4 = "Tiempos de disparo y cierre"
            T4F = "FR-0829"
        Case "Discrepancia de polos."
            T4 = "Discrepancia de polos"
            T4F = "???"
        Case "Test de verificación y operación Interruptor NSX."
            T4 = "Test de verificación y operación Interruptor NSX"
            T4F = "???"
        Case "Servicios Auxiliares."
            T4 = "Servicios Auxiliares"
            T4F = "???"
        Case "Dispositivos de supervisión y control"
            T4 = "Dispositivos de supervisión y control"
            T4F = "???"
    End Select
    
    ' Verificar que los combobox tienen valor - ENSAYO 5
    Select Case Me.OLEObjects("Ensayo5").Object.Value
        Case "Efectividad del circuito de protección."
            T5 = "Efectividad del circuito de protección"
            T5F = "FR-0804"
        Case "Espesor de pintura."
            T5 = "Espesor de pintura"
            T5F = "FR-0815"
        Case "Resistencia de aislamiento."
            T5 = "Resistencia de aislamiento"
            T5F = "FR-0811"
        Case "Operación y funcionamiento mecánico."
            T5 = "Operación y funcionamiento mecánico"
            T5F = "FR-0890"
        Case "Instrumentacion y medicion."
            T5 = "Instrumentacion y medicion"
            T5F = "???"
        Case "Conversión y distribución de energia."
            T5 = "Conversión y distribución de energia"
            T5F = "???"
        Case "Disposicion de supervisoon y control."
            T5 = "Disposicion de supervisoon y control"
            T5F = "???"
        Case "Inyección de corriente primaria por inducción (Circuito de medida.)"
            T5 = "Inyección de corriente primaria por inducción (Circuito de medida)"
            T5F = "???"
        Case "Tensión soportada a frecuencia industrial."
            T5 = "Tensión soportada a frecuencia industrial"
            T5F = "FR-0803"
        Case "Adherencia de pintura."
            T5 = "Adherencia de pintura"
            T5F = "???"
        Case "Grado de protección IP (Hasta 4X)."
            T5 = "Grado de protección IP (Hasta 4X)"
            T5F = "???"
        Case "Imborabilidad del Rotulado."
            T5 = "Imborabilidad del Rotulado"
            T5F = "???"
        Case "Permanecía del Rotulado."
            T5 = "Permanecía del Rotulado"
            T5F = "???"
        Case "Pruebas a unidad de disparo Micrologic SE."
            T5 = "Pruebas a unidad de disparo Micrologic SE"
            T5F = "FR-0829"
        Case "Tiempos de disparo y cierre."
            T5 = "Tiempos de disparo y cierre"
            T5F = "FR-0829"
        Case "Discrepancia de polos."
            T5 = "Discrepancia de polos"
            T5F = "???"
        Case "Test de verificación y operación Interruptor NSX."
            T5 = "Test de verificación y operación Interruptor NSX"
            T5F = "???"
        Case "Servicios Auxiliares."
            T5 = "Servicios Auxiliares"
            T5F = "???"
        Case "Dispositivos de supervisión y control"
            T5 = "Dispositivos de supervisión y control"
            T5F = "???"
    End Select
    
    ' Construir el texto de ensayos
    Ensayos_Personalizado = ""
    If T1 <> "" Then Ensayos_Personalizado = Ensayos_Personalizado & T1 & "/"
    If T2 <> "" Then Ensayos_Personalizado = Ensayos_Personalizado & T2 & "/"
    If T3 <> "" Then Ensayos_Personalizado = Ensayos_Personalizado & T3 & "/"
    If T4 <> "" Then Ensayos_Personalizado = Ensayos_Personalizado & T4 & "/"
    If T5 <> "" Then Ensayos_Personalizado = Ensayos_Personalizado & T5 & "/"
    
    ' Eliminar el último "/" si existe
    If Right(Ensayos_Personalizado, 1) = "/" Then
        Ensayos_Personalizado = Left(Ensayos_Personalizado, Len(Ensayos_Personalizado) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If Ensayos_Personalizado = "" Then
        MsgBox "No se ha seleccionado ningún ensayo Personalizado.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Solicitar la OT
    OT_Buscada = InputBox("Ingrese el número de OT donde desea cargar los ensayos Personalizados:", _
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
        ws.Cells(filaOT, "EF").Value = Ensayos_Personalizado
        
        MsgBox "Ensayos  cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Ensayos:" & vbCrLf & Ensayos_Personalizado, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " en la columna BA.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub

'-----------------------------------------------------------------------


