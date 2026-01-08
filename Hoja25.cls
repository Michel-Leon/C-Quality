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
    EQ1 = ""
    EQ2 = ""
    EQ3 = ""
    EQ4 = ""
    
    
    ' Si los CheckBox anteriores no funcionaron, intentar con OLEObjects (ActiveX Controls)
    If T1 = "" And T2 = "" And T3 = "" And T4 = "" And T5 = "" Then
        On Error Resume Next
        If ws.OLEObjects("SP1").Object.Value = True Then
            T1 = "Prueba a unidad de disparo Micrologic S.E"
            'AP1 = "FR-0829"
        End If
        If ws.OLEObjects("SP2").Object.Value = True Then
            T2 = "Dispositivo de protección contra sobretenciones (DPS)"
            'AP2 = "???"
        End If
        If ws.OLEObjects("SP3").Object.Value = True Then
            T3 = "Interruptores diferenciales"
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