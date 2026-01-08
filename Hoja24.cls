'============= HOJA DE APLICACIONES ========
'=======VARIABLES GLOBALES======
public Palabra As String' SE GUARDA EL TEXTO CORRESPONDIENTE


Public sub limpiar_tabla_APLICACIONES()
    'LIMPIAR VARIABLES
    AP1 = ""
    AP2 = ""
    AP3 = ""
    AP4 = ""
    AP5 = ""
    AP6 = ""
    AP7 = ""
    AP8 = ""

    ThisWorkbook.Sheets("APLICACIONES").range("BA84:EF140").Value = ""
end sub
     
public sub Worksheet_Activate()

    activeSheet.Shapes("Grupo I-MCC").Visible = False
    activeSheet.Shapes("Grupo L-Meter").Visible = False
    activeSheet.Shapes("Grupo O-Paralel").Visible = False
    activeSheet.Shapes("Grupo P-ATS").Visible = False
    activeSheet.Shapes("Grupo Q-Bank").Visible = False
    activeSheet.Shapes("Grupo V-Drive").Visible = False
    ActiveSheet.PAST31.Enabled = False
    ActiveSheet.PAST32.Enabled = False
    ActiveSheet.PAST33.Enabled = False
    ActiveSheet.PAST34.Enabled = False
    ActiveSheet.PAST35.Enabled = False

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
    
    ' Inicializar variables
    T1 = ""
    T2 = ""
    T3 = ""
    T4 = "" 
    T5 = ""
    AP1 = ""
    AP2 = ""
    AP3 = ""
    AP4 = ""
    AP5 = ""
    AP6 = ""
    AP7 = ""
    AP8 = ""
    
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("IMCC1").Object.Value = True Then
            T1 = "Tableros para distribución Incomer & feeders (D)"
            'AP1 = "FR-0806"
        End If
        If ws.OLEObjects("IMCC2").Object.Value = True Then
            T2 = "Tab para alimentadores de potencia y arranque motor extraibles (W)"
            'AP2 = "???"
        End If
        If ws.OLEObjects("IMCC3").Object.Value = True Then
            T3 = "Tab para variadores de velocidad y arrancadores suaves (MS)"
            'AP3 = "FR-0811"
        End If
        If ws.OLEObjects("IMCC4").Object.Value = True Then
            T4 = "Tab para alimentadores y arranque motor fijas (MF)"
            'AP4 = "???"
        End If
        If ws.OLEObjects("IMCC5").Object.Value = True Then
            T5 = "T. para corrección de factor de potencia y corrección de armonicos (PQ)"
            'AP4 = "???"
        End If
        On Error GoTo 0
    End If
    
    ' Construir el texto de IMCC
    IMCC = ""
    If T1 <> "" Then IMCC = IMCC & T1 & "/"
    If T2 <> "" Then IMCC = IMCC & T2 & "/"
    If T3 <> "" Then IMCC = IMCC & T3 & "/"
    If T4 <> "" Then IMCC = IMCC & T4 & "/"
    If T5 <> "" Then IMCC = IMCC & T5 & "/"
    
    ' Eliminar el último "/" si existe
    If Right(IMCC, 1) = "/" Then
        IMCC = Left(IMCC, Len(IMCC) - 1)
    End If
    
    ' Verificar que hay formatos seleccionados
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
        ws.Cells(filaOT, "EF").Value = IMCC
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & IMCC, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & ". Verifique la OT.", _
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
    ' Inicializar variables
        T1 = ""
        T2 = ""
        T3 = ""
        T4 = "" 
        T5 = ""
        AP1 = ""
        AP2 = ""
        AP3 = ""
        AP4 = ""
        AP5 = ""
        AP6 = ""
        AP7 = ""
        AP8 = ""
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.LMeter1.Value = True Then
        T1 = "Tableros para medidores de energía ( D - S.D - P.M )"
        'AP1 = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
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
    ' Inicializar variables
        T1 = ""
        T2 = ""
        T3 = ""
        T4 = "" 
        T5 = ""
        AP1 = ""
        AP2 = ""
        AP3 = ""
        AP4 = ""
        AP5 = ""
        AP6 = ""
        AP7 = ""
        AP8 = ""
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.OParalel1.Value = True Then
        T1 = "Tableros de sincronismo"
        'AP1 = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
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
' encender check para P-ATS ASCO
Private Sub PAST3_Click()
    IF ActiveSheet.PAST3.Value = True Then
        ActiveSheet.PAST31.Enabled = True
        ActiveSheet.PAST32.Enabled = True
        ActiveSheet.PAST33.Enabled = True
        ActiveSheet.PAST34.Enabled = True
        ActiveSheet.PAST35.Enabled = True
    Else
        ActiveSheet.PAST31.Enabled = False
        ActiveSheet.PAST32.Enabled = False
        ActiveSheet.PAST33.Enabled = False
        ActiveSheet.PAST34.Enabled = False
        ActiveSheet.PAST35.Enabled = False
    End If    
End Sub
Public Sub Cargar_PATS()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String, T6 As String, T7 As String, T8 As String
    Dim PATS As String
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
        AP1 = ""
        AP2 = ""
        AP3 = ""
        AP4 = ""
        AP5 = ""
        AP6 = ""
        AP7 = ""
        AP8 = ""
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.PATS1.Value = True Then
        T1 = "ATS transición abierta"
        'AP1 = "FR-0805" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    If ActiveSheet.PATS2.Value = True Then
        T2 = "ATS transición cerrada"
        'AP2 = "FR-0805" 'no se encontro referencia del formato
    End If
    If ActiveSheet.PATS3.Value = True Then
        T3 = "ATS ASCO  (Suministro fabricación SCHNEIDER)"
        'AP3 = "FR-0805" 'no se encontro referencia del formato
    End If
    If ActiveSheet.PAST31.Value = True Then
        T4 = "Verificación mecánica"
        AP4 = "FR-0833" ' cARGADO
    End If
    IF ActiveSheet.PAST32.Value = True Then
        T5 = "Verificación Barraje"
        AP5 = "FR-0834" ' cARGADO
    end If
    if ActiveSheet.PAST33.Value = True Then
        T6 = "Verificación de cableado"
        AP6 = "FR-0835" ' cARGADO
    end if
    if ActiveSheet.PAST34.Value = True Then
        T7 = "Automatic transfer switch grupo 5"
        'AP7= "FR-0836" ' no se encontro referencia del formato
    end if
    if ActiveSheet.PAST35.Value = True Then
        T8 = "Automatic transfer switch grupo G"
        'AP8 = "FR-0837" ' no se encontro referencia del formato
    end if
    ' Construir el texto de ensayos (eliminar líneas vacías)
    PATS = ""
    If T1 <> "" Then PATS = PATS & T1 & "/"
    If T2 <> "" Then PATS = PATS & T2 & "/"
    If T3 <> "" Then PATS = PATS & T3 & "("
    If T4 <> "" Then PATS = PATS & T4 & "/"
    If T5 <> "" Then PATS = PATS & T5 & "/"
    If T6 <> "" Then PATS = PATS & T6 & "/"
    If T7 <> "" Then PATS = PATS & T7 & "/"
    If T8 <> "" Then PATS = PATS & T8 & ")"
    
    ' Limpiar saltos de línea finales
    PATS = Trim(PATS)
    If Right(PATS, 1) = Chr(10) Then
        PATS = Left(PATS, Len(PATS) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If PATS = "" Then
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
        ws.Cells(filaOT, "EF").Value = PATS
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & PATS, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_QBANK()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim QBANK As String
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
        AP1 = ""
        AP2 = ""
        AP3 = ""
        AP4 = ""
        AP5 = ""
        AP6 = ""
        AP7 = ""
        AP8 = ""
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.QBANK1.Value = True Then
        T1 = "Tablero para banco de condensadores (PF)"
        'AP1 = "FR-0814" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    If ActiveSheet.QBANK2.Value = True Then
        T2 = "Tablero para filtros pasivos (HF)"
        'AP2 = "???" 'no se encontro referencia del formato
    End If
    If ActiveSheet.QBANK3.Value = True Then
        T3 = "Tablero para banco de cargas resisitivos (RL)"
        'AP3 = "FR-???"'cargado
    End If
    If ActiveSheet.QBANK4.Value = True Then
        T4 = "Compensadores dinámicos (CD)"
        'AP4 = "???" 'no se encontro referencia del formato
    End If
    
    ' Construir el texto de ensayos (eliminar líneas vacías)
    QBANK = ""
    If T1 <> "" Then QBANK = QBANK & T1 & "/"
    If T2 <> "" Then QBANK = QBANK & T2 & "/"
    If T3 <> "" Then QBANK = QBANK & T3 & "/"
    If T4 <> "" Then QBANK = QBANK & T4 & "/"
    If T5 <> "" Then QBANK = QBANK & T5
    
    ' Limpiar saltos de línea finales
    QBANK = Trim(QBANK)
    If Right(QBANK, 1) = Chr(10) Then
        QBANK = Left(QBANK, Len(QBANK) - 1)
    End If
    
    ' Verificar que hay ensayos seleccionados
    If QBANK = "" Then
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
        ws.Cells(filaOT, "EF").Value = QBANK
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & QBANK, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
    
End Sub
Public Sub Cargar_VDRIVE()
    Dim T1 As String, T2 As String, T3 As String, T4 As String, T5 As String
    Dim VDRIVE As String
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
        AP1 = ""
        AP2 = ""
        AP3 = ""
        AP4 = ""
        AP5 = ""
        AP6 = ""
        AP7 = ""
        AP8 = ""
    
    ' Verificar que los CheckBox tienen valor
    If ActiveSheet.VDRIVE1.Value = True Then
        T1 = "Tableros para variadores de velocidad"
        'T1F = "???" 'NO SE ENCONTRO REFERENCIA DEL FORMATO
    End If
    ' Construir el texto de ensayos (eliminar líneas vacías)
    VDRIVE = ""
    If T1 <> "" Then VDRIVE = VDRIVE & T1
    ' Limpiar saltos de línea finales
    VDRIVE = Trim(VDRIVE)
    If Right(VDRIVE, 1) = Chr(10) Then
        VDRIVE = Left(VDRIVE, Len(VDRIVE) - 1)
    End If
    ' Verificar que hay ensayos seleccionados
    If VDRIVE = "" Then
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
        ws.Cells(filaOT, "EF").Value = VDRIVE
        
        MsgBox "Formatos Cargados correctamente en la OT: " & OT_Buscada & _
               vbCrLf & "Fila: " & filaOT & vbCrLf & vbCrLf & _
               "Formatos de la aplicación:" & vbCrLf & VDRIVE, vbInformation, "Éxito"
    Else
        MsgBox "No se encontró la OT: " & OT_Buscada & " Verifique la OT.", _
               vbExclamation, "OT no encontrada"
    End If
End Sub