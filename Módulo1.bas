Attribute VB_Name = "M�dulo1"
'=======VARIABLES GLOBALES======
public nombreCarpetaProyecto As String ' varible global
public Cliente as string
public Elaborado_por as string
Public Producto As String
Public Aplicacion As String
PUBLIC OT_General As String
public OT_Buscar As String
public T1F As String, T2F As String, T3F As String, T4F As String, T5F As String ' para ensayos segun el tipo
public EF1 AS STRING, EF2 AS STRING, EF3 AS STRING, EF4 AS STRING, EF5 AS STRING, EF6 AS STRING ' para ensayos FIJOS SEGUN EL TIPO
PUBLIC EF7 AS STRING, EF8 AS STRING, EF9 AS STRING, EF10 AS STRING, EF11 AS STRING, EF12 AS STRING, EF13 AS STRING ' para ensayos FIJOS SEGUN EL TIPO
Public AP1 AS STRING, AP2 AS STRING, AP3 AS STRING, AP4 AS STRING, AP5 AS STRING, AP6 AS STRING, AP7 AS STRING, AP8 AS STRING ' PARA FORMATOS DE APLICAICONES
Public EQ1 AS STRING, EQ2 AS STRING, EQ3 AS STRING, EQ4 AS STRING ' PARA FORMATOS DE EQUIPOS

'TOMARR VALORES DE LAS CASILLAS
Public CH1 as string, CH2 as string, CH3 as string, CH4 as string, CH5 as string, CH6 as string
public columna1 as string, columna2 as string, columna3 as string, columna4 as string, columna5 as string, columna6 as string
Sub Tomar_valores()
    Dim HOJA_INSPECCION As Worksheet
    Dim filaActual As Long
    Dim encontrado As Boolean
    
    Set HOJA_INSPECCION = ThisWorkbook.Worksheets("INSPECCION")
    
    ' Solicitar la OT al usuario
    OT_Buscar = InputBox("Ingrese el número de OT a buscar:", "Buscar OT")
    
    ' Verificar si el usuario canceló o no ingresó nada
    If OT_Buscar = "" Then
        MsgBox "No se ingresó ninguna OT. Operación cancelada.", vbExclamation
        Exit Sub
    End If
    
    ' Buscar la OT en la columna BA desde la fila 84, avanzando de 4 en 4
    encontrado = False
    filaActual = 84
    
    Do While HOJA_INSPECCION.Range("BA" & filaActual).Value <> ""
        If HOJA_INSPECCION.Range("BA" & filaActual).Value = OT_Buscar Then
            ' Se encontró la OT, tomar los valores de esa fila
            CH1 = HOJA_INSPECCION.Range("EO" & filaActual).Value
            CH2 = HOJA_INSPECCION.Range("FA" & filaActual).Value
            CH3 = HOJA_INSPECCION.Range("FY" & filaActual).Value
            CH4 = HOJA_INSPECCION.Range("GW" & filaActual).Value
            CH5 = HOJA_INSPECCION.Range("HU" & filaActual).Value
            CH6 = HOJA_INSPECCION.Range("IS" & filaActual).Value
            columna1 = HOJA_INSPECCION.Range("EF" & filaActual).Value
            columna2 = HOJA_INSPECCION.Range("ER" & filaActual).Value
            columna3 = HOJA_INSPECCION.Range("FD" & filaActual).Value
            columna4 = HOJA_INSPECCION.Range("GB" & filaActual).Value
            columna5 = HOJA_INSPECCION.Range("GZ" & filaActual).Value
            columna6 = HOJA_INSPECCION.Range("HX" & filaActual).Value
            
            encontrado = True
            'MsgBox "OT encontrada en la fila " & filaActual & vbCrLf & _
                   "Valores capturados correctamente.", vbInformation
            Exit Do
        End If
        
        filaActual = filaActual + 4
        
        ' Opcional: limitar la búsqueda a un rango razonable (por ejemplo, 1000 filas)
        If filaActual > 10000 Then
            Exit Do
        End If
    Loop
    
    ' Si no se encontró la OT
    If Not encontrado Then
        MsgBox "No se encontró la OT: " & OT_Buscar, vbExclamation
    End If
End Sub

sub leer()
    CH1 = ThisWorkbook.Worksheets("INSPECCION").Range("EO84").Value
    MsgBox CH1
end sub

Public Sub crearPaqueteAseguramiento()
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim hojascopiar As Collection
    Dim hoja As Worksheet
    Dim nombreHoja As Variant
    Dim rutaBase As String
    Dim nombreCarpetaPrincipal As String
    Dim nombreCarpetaAseguramiento As String
    Dim rutaCarpetaProyecto As String
    Dim rutaCarpetaPrincipal As String
    Dim rutaCarpetaAseguramiento As String
    Dim nombreArchivo As String
    Dim estadoVerificacion As String
    Dim HOJA_INSPECCION As Worksheet
    Dim hoja_Proyecto As Worksheet
    Dim filaOT_Inspeccion As Long
    Dim filaOT_Proyecto As Long
    Dim fechaCreacion As String
    Dim fso As Object
    Dim encontradoInspeccion As Boolean
    Dim encontradoProyecto As Boolean
    Dim rutaTemporal As String
    Dim rutaCompleta As String
    
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection
    Set HOJA_INSPECCION = wbOrigen.Worksheets("INSPECCION")
    Set hoja_Proyecto = wbOrigen.Worksheets("PROYECTO")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la fecha actual en formato YYYY-MM-DD
    fechaCreacion = Format(Date, "yyyy-mm-dd")
    
    ' ========== OBTENER NOMBRE DEL PROYECTO DESDE CT59 ==========
    nombreCarpetaProyecto = Trim(wbOrigen.Worksheets("PROYECTO").Range("CT59").Value)
    
    ' Validar que el nombre del proyecto no esté vacío
    If nombreCarpetaProyecto = "" Then
        MsgBox "No se ha ingresado un nombre de proyecto." & vbCrLf & _
               "Por favor, ingrese un nombre de proyecto antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar caracteres inválidos del nombre de carpeta proyecto
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "/", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "\", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ":", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "*", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "?", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, """", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "<", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ">", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "|", "-")
    
    ' ========== BUSCAR OT EN HOJA INSPECCION ==========
    encontradoInspeccion = False
    filaOT_Inspeccion = 84
    Do While HOJA_INSPECCION.Range("BA" & filaOT_Inspeccion).Value <> ""
        If HOJA_INSPECCION.Range("BA" & filaOT_Inspeccion).Value = OT_Buscar Then
            ' Obtener valores de los checkboxes de la fila encontrada
            CH1 = HOJA_INSPECCION.Range("EO" & filaOT_Inspeccion).Value
            CH2 = HOJA_INSPECCION.Range("FA" & filaOT_Inspeccion).Value
            CH3 = HOJA_INSPECCION.Range("FY" & filaOT_Inspeccion).Value
            CH4 = HOJA_INSPECCION.Range("GW" & filaOT_Inspeccion).Value
            CH5 = HOJA_INSPECCION.Range("HU" & filaOT_Inspeccion).Value
            CH6 = HOJA_INSPECCION.Range("IS" & filaOT_Inspeccion).Value
            
            encontradoInspeccion = True
            Exit Do
        End If
        filaOT_Inspeccion = filaOT_Inspeccion + 4
        If filaOT_Inspeccion > 10000 Then Exit Do
    Loop
    
    ' Verificar si se encontró en INSPECCION
    If Not encontradoInspeccion Then
        MsgBox "No se encontró la OT " & OT_Buscar & " en la hoja INSPECCION.", vbExclamation
        Exit Sub
    End If
    
    ' ========== BUSCAR OT EN HOJA PROYECTO ==========
    encontradoProyecto = False
    filaOT_Proyecto = 96
    Do While hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value <> ""
        If hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value = OT_Buscar Then
            ' Obtener el estado de verificación de la columna EP
            estadoVerificacion = hoja_Proyecto.Range("EP" & filaOT_Proyecto).Value
            encontradoProyecto = True
            Exit Do
        End If
        filaOT_Proyecto = filaOT_Proyecto + 4
        If filaOT_Proyecto > 10000 Then Exit Do
    Loop
    
    ' Verificar si se encontró en PROYECTO
    If Not encontradoProyecto Then
        MsgBox "No se encontró la OT " & OT_Buscar & " en la hoja PROYECTO.", vbExclamation
        Exit Sub
    End If
    
    ' ========== VERIFICAR ESTADO ==========
    If UCase(Trim(estadoVerificacion)) <> "VERIFICADO" Then
        MsgBox "La OT " & OT_Buscar & " no ha sido verificada." & vbCrLf & _
               "Estado actual: " & estadoVerificacion & vbCrLf & _
               "No se puede crear el paquete de aseguramiento.", vbExclamation
        Exit Sub
    End If
    
    ' ========== AGREGAR HOJAS SEGÚN LOS CHECKBOXES ==========
    Select Case CH1
        Case True, "Verdadero", "VERDADERO"
            Select Case columna1
                Case "FR-0856"
                    If SheetExists("FR-0856") Then hojascopiar.Add wbOrigen.Worksheets("FR-0856")
                Case "FR-0850"
                    If SheetExists("FR-0850") Then hojascopiar.Add wbOrigen.Worksheets("FR-0850")
            End Select
    End Select
    Select Case CH2
        Case True, "Verdadero", "VERDADERO"
            Select Case columna2
                Case "FR-0857"
                    If SheetExists("FR-0857") Then hojascopiar.Add wbOrigen.Worksheets("FR-0857")
                Case "FR-0851"
                    If SheetExists("FR-0851") Then hojascopiar.Add wbOrigen.Worksheets("FR-0851")
            End Select
    End Select
    Select Case CH3
        Case True, "Verdadero", "VERDADERO"
            Select Case columna3
                Case "FR-0858"
                    If SheetExists("FR-0858") Then hojascopiar.Add wbOrigen.Worksheets("FR-0858")
                Case "FR-0852"
                    If SheetExists("FR-0852") Then hojascopiar.Add wbOrigen.Worksheets("FR-0852")
            End Select
    End Select
    Select Case CH4
        Case True, "Verdadero", "VERDADERO"
            Select Case columna4
                Case "FR-0859"
                    If SheetExists("FR-0859") Then hojascopiar.Add wbOrigen.Worksheets("FR-0859")
                Case "FR-0853"
                    If SheetExists("FR-0853") Then hojascopiar.Add wbOrigen.Worksheets("FR-0853")
            End Select
    End Select
    Select Case CH5
        Case True, "Verdadero", "VERDADERO"
            Select Case columna5
                Case "FR-0860"
                    If SheetExists("FR-0860") Then hojascopiar.Add wbOrigen.Worksheets("FR-0860")
                Case "FR-0854"
                    If SheetExists("FR-0854") Then hojascopiar.Add wbOrigen.Worksheets("FR-0854")
            End Select
    End Select
    Select Case CH6
        Case True, "Verdadero", "VERDADERO"
            Select Case columna6
                Case "FR-0861"
                    If SheetExists("FR-0861") Then hojascopiar.Add wbOrigen.Worksheets("FR-0861")
                Case "FR-0855"
                    If SheetExists("FR-0855") Then hojascopiar.Add wbOrigen.Worksheets("FR-0855")
            End Select
    End Select
    
    ' Verificar si hay hojas para copiar
    If hojascopiar.Count = 0 Then
        MsgBox "No hay hojas seleccionadas para copiar." & vbCrLf & _
               "Verifique que los checkboxes estén marcados en la hoja INSPECCION.", vbExclamation
        Exit Sub
    End If
    
    ' ========== CREAR ESTRUCTURA DE CARPETAS ==========
    ' Definir ruta base
    rutaBase = "C:\Users\sleon\OneDrive - industriascts.com\Pruebas FAT\"
    
    ' Crear carpeta de proyecto
    rutaCarpetaProyecto = rutaBase & nombreCarpetaProyecto & "\"
    If Not fso.FolderExists(rutaCarpetaProyecto) Then
        fso.CreateFolder rutaCarpetaProyecto
    End If
    
    ' Crear nombre de carpeta principal (OT): OT-Producto-Aplicacion-Fecha
    nombreCarpetaPrincipal = "OT-" & OT_Buscar & "-" & Producto & "-" & Aplicacion & "-" & fechaCreacion
    rutaCarpetaPrincipal = rutaCarpetaProyecto & nombreCarpetaPrincipal & "\"
    
    ' Crear carpeta de OT si no existe
    If Not fso.FolderExists(rutaCarpetaPrincipal) Then
        fso.CreateFolder rutaCarpetaPrincipal
    End If
    
    ' Crear carpeta "01.Inspeccion y verificacion"
    nombreCarpetaAseguramiento = "01.Inspeccion y verificacion"
    rutaCarpetaAseguramiento = rutaCarpetaPrincipal & nombreCarpetaAseguramiento & "\"
    
    If Not fso.FolderExists(rutaCarpetaAseguramiento) Then
        fso.CreateFolder rutaCarpetaAseguramiento
    End If
    
    ' ========== CREAR Y GUARDAR LIBRO DE EXCEL ==========
    ' Crear nuevo libro de Excel
    Set wbNuevo = Workbooks.Add
    
    ' Eliminar hojas predeterminadas del nuevo libro (excepto una)
    Application.DisplayAlerts = False
    Do While wbNuevo.Worksheets.Count > 1
        wbNuevo.Worksheets(wbNuevo.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Copiar las hojas seleccionadas al nuevo libro
    For Each hoja In hojascopiar
        hoja.Copy After:=wbNuevo.Worksheets(wbNuevo.Worksheets.Count)
    Next hoja
    
    ' Eliminar la hoja vacía inicial
    Application.DisplayAlerts = False
    If wbNuevo.Worksheets.Count > 1 Then
        wbNuevo.Worksheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Preparar nombre de archivo
    nombreArchivo = OT_Buscar & "-" & "Paquete de Aseguramiento.xlsx"
    
    ' Limpiar caracteres inválidos del nombre de archivo
    nombreArchivo = Replace(nombreArchivo, "/", "-")
    nombreArchivo = Replace(nombreArchivo, "\", "-")
    nombreArchivo = Replace(nombreArchivo, ":", "-")
    nombreArchivo = Replace(nombreArchivo, "*", "-")
    nombreArchivo = Replace(nombreArchivo, "?", "-")
    nombreArchivo = Replace(nombreArchivo, """", "-")
    nombreArchivo = Replace(nombreArchivo, "<", "-")
    nombreArchivo = Replace(nombreArchivo, ">", "-")
    nombreArchivo = Replace(nombreArchivo, "|", "-")
    
    ' Ruta completa del archivo final
    rutaCompleta = rutaCarpetaAseguramiento & nombreArchivo
    
    ' Ruta temporal para guardar primero
    rutaTemporal = Environ("TEMP") & "\" & nombreArchivo
    
    ' Eliminar archivo temporal si ya existe
    On Error Resume Next
    If Dir(rutaTemporal) <> "" Then Kill rutaTemporal
    On Error GoTo 0
    
    ' Guardar en ubicación temporal
    On Error GoTo ErrorGuardar
    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Eliminar archivo en destino final si existe
    On Error Resume Next
    If Dir(rutaCompleta) <> "" Then Kill rutaCompleta
    On Error GoTo 0
    
    ' Mover el archivo de temporal a ubicación final
    On Error GoTo ErrorMover
    fso.MoveFile rutaTemporal, rutaCompleta
    On Error GoTo 0
    
    ' ========== MENSAJE FINAL ==========
    MsgBox "Paquete de aseguramiento creado exitosamente en:" & vbCrLf & _
           rutaCarpetaAseguramiento & nombreArchivo & vbCrLf & vbCrLf & _
           "Proyecto: " & nombreCarpetaProyecto & vbCrLf & _
           "Hojas copiadas: " & hojascopiar.Count, vbInformation, "Proceso completado"
    
    ' Abrir la carpeta
    Shell "explorer.exe """ & rutaCarpetaAseguramiento & """", vbNormalFocus
    
    Exit Sub

    ErrorGuardar:
        Application.DisplayAlerts = True
        MsgBox "Error al guardar el archivo temporal:" & vbCrLf & _
            "Ruta: " & rutaTemporal & vbCrLf & vbCrLf & _
            "Error: " & Err.Description, vbCritical
        
        On Error Resume Next
        wbNuevo.Close SaveChanges:=False
        Exit Sub

    ErrorMover:
        MsgBox "Error al mover el archivo a la ubicación final:" & vbCrLf & _
            "Desde: " & rutaTemporal & vbCrLf & _
            "Hacia: " & rutaCompleta & vbCrLf & vbCrLf & _
            "Error: " & Err.Description & vbCrLf & vbCrLf & _
            "El archivo se guardó en la carpeta temporal." & vbCrLf & _
            "Puede moverlo manualmente.", vbCritical
        
        ' Abrir carpeta temporal
        Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Exit Sub
End Sub

'Crear paquete de ensayos dentro de la carpeta global de proyecto
Public Sub crearPaqueteEnsayos()
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim hojascopiar As Collection
    Dim hoja As Worksheet
    Dim rutaBase As String
    Dim nombreCarpetaProyecto As String
    Dim nombreCarpetaPrincipal As String
    Dim nombreCarpetaEnsayos As String
    Dim rutaCarpetaProyecto As String
    Dim rutaCarpetaPrincipal As String
    Dim rutaCarpetaEnsayos As String
    Dim nombreArchivo As String
    Dim estadoVerificacionEnsayos As String
    Dim hoja_Proyecto As Worksheet
    Dim filaOT_Proyecto As Long
    Dim fechaCreacion As String
    Dim fso As Object
    Dim encontradoProyecto As Boolean
    Dim rutaTemporal As String
    Dim rutaCompleta As String
    
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection
    Set hoja_Proyecto = wbOrigen.Worksheets("PROYECTO")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la fecha actual en formato YYYY-MM-DD
    fechaCreacion = Format(Date, "yyyy-mm-dd")
    
    ' ========== OBTENER NOMBRE DEL PROYECTO DESDE CT59 ==========
    nombreCarpetaProyecto = Trim(wbOrigen.Worksheets("PROYECTO").Range("CT59").Value)
    
    ' Validar que el nombre del proyecto no esté vacío
    If nombreCarpetaProyecto = "" Then
        MsgBox "La celda CT59 no contiene un nombre de proyecto." & vbCrLf & _
               "Por favor, ingrese un nombre de proyecto antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar caracteres inválidos del nombre de carpeta proyecto
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "/", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "\", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ":", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "*", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "?", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, """", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "<", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ">", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "|", "-")
    
    ' ========== BUSCAR OT EN HOJA PROYECTO ==========
    encontradoProyecto = False
    filaOT_Proyecto = 96
    Do While hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value <> ""
        If hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value = OT_Buscar Then
            ' Obtener el estado de verificación de ensayos de la columna FP
            estadoVerificacionEnsayos = hoja_Proyecto.Range("FP" & filaOT_Proyecto).Value
            encontradoProyecto = True
            Exit Do
        End If
        filaOT_Proyecto = filaOT_Proyecto + 4
        If filaOT_Proyecto > 10000 Then Exit Do
    Loop
    
    ' Verificar si se encontró en PROYECTO
    If Not encontradoProyecto Then
        MsgBox "No se encontró la OT " & OT_Buscar & " en la hoja PROYECTO.", vbExclamation
        Exit Sub
    End If
    
    ' ========== VERIFICAR ESTADO DE ENSAYOS ==========
    If UCase(Trim(estadoVerificacionEnsayos)) <> "VERIFICADO" Then
        MsgBox "La OT " & OT_Buscar & " no tiene los ensayos verificados." & vbCrLf & _
               "Estado actual en columna FP: " & estadoVerificacionEnsayos & vbCrLf & _
               "No se puede crear el paquete de ensayos.", vbExclamation
        Exit Sub
    End If
    
    ' ========== AGREGAR HOJAS SEGÚN LAS VARIABLES T1F-T5F ==========
    ' Agregar hoja T1F si tiene valor
    If Trim(T1F) <> "" Then
        If SheetExists(T1F) Then
            hojascopiar.Add wbOrigen.Worksheets(T1F)
        Else
            MsgBox "Advertencia: La hoja '" & T1F & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja T2F si tiene valor
    If Trim(T2F) <> "" Then
        If SheetExists(T2F) Then
            hojascopiar.Add wbOrigen.Worksheets(T2F)
        Else
            MsgBox "Advertencia: La hoja '" & T2F & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja T3F si tiene valor
    If Trim(T3F) <> "" Then
        If SheetExists(T3F) Then
            hojascopiar.Add wbOrigen.Worksheets(T3F)
        Else
            MsgBox "Advertencia: La hoja '" & T3F & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja T4F si tiene valor
    If Trim(T4F) <> "" Then
        If SheetExists(T4F) Then
            hojascopiar.Add wbOrigen.Worksheets(T4F)
        Else
            MsgBox "Advertencia: La hoja '" & T4F & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja T5F si tiene valor
    If Trim(T5F) <> "" Then
        If SheetExists(T5F) Then
            hojascopiar.Add wbOrigen.Worksheets(T5F)
        Else
            MsgBox "Advertencia: La hoja '" & T5F & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Verificar si hay hojas para copiar
    If hojascopiar.Count = 0 Then
        MsgBox "No hay hojas de ensayos para copiar." & vbCrLf & _
               "Verifique que las variables T1F a T5F contengan nombres de hojas válidos.", vbExclamation
        Exit Sub
    End If
    
    ' ========== CREAR ESTRUCTURA DE CARPETAS ==========
    ' Definir ruta base
    rutaBase = "C:\Users\sleon\OneDrive - industriascts.com\Pruebas FAT\"
    
    ' Ruta de la carpeta del proyecto
    rutaCarpetaProyecto = rutaBase & nombreCarpetaProyecto & "\"
    If Not fso.FolderExists(rutaCarpetaProyecto) Then
        fso.CreateFolder rutaCarpetaProyecto
    End If
    
    ' Crear nombre de carpeta principal (OT): OT-Producto-Aplicacion-Fecha
    nombreCarpetaPrincipal = "OT-" & OT_Buscar & "-" & Producto & "-" & Aplicacion & "-" & fechaCreacion
    rutaCarpetaPrincipal = rutaCarpetaProyecto & nombreCarpetaPrincipal & "\"
    
    ' Crear carpeta de OT si no existe
    If Not fso.FolderExists(rutaCarpetaPrincipal) Then
        fso.CreateFolder rutaCarpetaPrincipal
    End If
    
    ' Crear carpeta "02.Ensayos"
    nombreCarpetaEnsayos = "02.Ensayos"
    rutaCarpetaEnsayos = rutaCarpetaPrincipal & nombreCarpetaEnsayos & "\"
    
    If Not fso.FolderExists(rutaCarpetaEnsayos) Then
        fso.CreateFolder rutaCarpetaEnsayos
    End If
    
    ' ========== CREAR Y GUARDAR LIBRO DE EXCEL ==========
    ' Crear nuevo libro de Excel
    Set wbNuevo = Workbooks.Add
    
    ' Eliminar hojas predeterminadas del nuevo libro (excepto una)
    Application.DisplayAlerts = False
    Do While wbNuevo.Worksheets.Count > 1
        wbNuevo.Worksheets(wbNuevo.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Copiar las hojas seleccionadas al nuevo libro
    For Each hoja In hojascopiar
        hoja.Copy After:=wbNuevo.Worksheets(wbNuevo.Worksheets.Count)
    Next hoja
    
    ' Eliminar la hoja vacía inicial
    Application.DisplayAlerts = False
    If wbNuevo.Worksheets.Count > 1 Then
        wbNuevo.Worksheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Preparar nombre de archivo
    nombreArchivo = OT_Buscar & "-" & "Paquete de Ensayos.xlsx"
    
    ' Limpiar caracteres inválidos del nombre de archivo
    nombreArchivo = Replace(nombreArchivo, "/", "-")
    nombreArchivo = Replace(nombreArchivo, "\", "-")
    nombreArchivo = Replace(nombreArchivo, ":", "-")
    nombreArchivo = Replace(nombreArchivo, "*", "-")
    nombreArchivo = Replace(nombreArchivo, "?", "-")
    nombreArchivo = Replace(nombreArchivo, """", "-")
    nombreArchivo = Replace(nombreArchivo, "<", "-")
    nombreArchivo = Replace(nombreArchivo, ">", "-")
    nombreArchivo = Replace(nombreArchivo, "|", "-")
    
    ' Ruta completa del archivo final
    rutaCompleta = rutaCarpetaEnsayos & nombreArchivo
    
    ' Ruta temporal para guardar primero
    rutaTemporal = Environ("TEMP") & "\" & nombreArchivo
    
    ' Eliminar archivo temporal si ya existe
    On Error Resume Next
    If Dir(rutaTemporal) <> "" Then Kill rutaTemporal
    On Error GoTo 0
    
    ' Guardar en ubicación temporal
    On Error GoTo ErrorGuardar
    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Eliminar archivo en destino final si existe
    On Error Resume Next
    If Dir(rutaCompleta) <> "" Then Kill rutaCompleta
    On Error GoTo 0
    
    ' Mover el archivo de temporal a ubicación final
    On Error GoTo ErrorMover
    fso.MoveFile rutaTemporal, rutaCompleta
    On Error GoTo 0
    
    ' ========== MENSAJE FINAL ==========
    MsgBox "Paquete de ensayos creado exitosamente en:" & vbCrLf & _
           rutaCarpetaEnsayos & nombreArchivo & vbCrLf & vbCrLf & _
           "Proyecto: " & nombreCarpetaProyecto & vbCrLf & _
           "Hojas copiadas: " & hojascopiar.Count, vbInformation, "Proceso completado"
    
    ' Abrir la carpeta
    Shell "explorer.exe """ & rutaCarpetaEnsayos & """", vbNormalFocus
    
    Exit Sub

ErrorGuardar:
    Application.DisplayAlerts = True
    MsgBox "Error al guardar el archivo temporal:" & vbCrLf & _
           "Ruta: " & rutaTemporal & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical
    
    On Error Resume Next
    wbNuevo.Close SaveChanges:=False
    Exit Sub

ErrorMover:
    MsgBox "Error al mover el archivo a la ubicación final:" & vbCrLf & _
           "Desde: " & rutaTemporal & vbCrLf & _
           "Hacia: " & rutaCompleta & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "El archivo se guardó en la carpeta temporal." & vbCrLf & _
           "Puede moverlo manualmente.", vbCritical
    
    ' Abrir carpeta temporal
    Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
    Exit Sub
End Sub
Public Sub crearPaqueteAplicaciones()
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim hojascopiar As Collection
    Dim hoja As Worksheet
    Dim rutaBase As String
    Dim nombreCarpetaProyecto As String
    Dim nombreCarpetaPrincipal As String
    Dim nombreCarpetaAplicaciones As String
    Dim rutaCarpetaProyecto As String
    Dim rutaCarpetaPrincipal As String
    Dim rutaCarpetaAplicaciones As String
    Dim nombreArchivo As String
    Dim estadoVerificacionAplicaciones As String
    Dim hoja_Proyecto As Worksheet
    Dim filaOT_Proyecto As Long
    Dim fechaCreacion As String
    Dim fso As Object
    Dim encontradoProyecto As Boolean
    Dim rutaTemporal As String
    Dim rutaCompleta As String
    
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection
    Set hoja_Proyecto = wbOrigen.Worksheets("PROYECTO")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la fecha actual en formato YYYY-MM-DD
    fechaCreacion = Format(Date, "yyyy-mm-dd")
    
    ' ========== OBTENER NOMBRE DEL PROYECTO DESDE CT59 ==========
    nombreCarpetaProyecto = Trim(wbOrigen.Worksheets("PROYECTO").Range("CT59").Value)
    
    ' Validar que el nombre del proyecto no esté vacío
    If nombreCarpetaProyecto = "" Then
        MsgBox "La celda CT59 no contiene un nombre de proyecto." & vbCrLf & _
               "Por favor, ingrese un nombre de proyecto antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar caracteres inválidos del nombre de carpeta proyecto
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "/", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "\", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ":", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "*", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "?", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, """", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "<", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ">", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "|", "-")
    
    ' ========== BUSCAR OT EN HOJA PROYECTO ==========
    encontradoProyecto = False
    filaOT_Proyecto = 96
    Do While hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value <> ""
        If hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value = OT_Buscar Then
            ' Obtener el estado de verificación de ensayos de la columna FP
            estadoVerificacionAplicaciones = hoja_Proyecto.Range("FP" & filaOT_Proyecto).Value
            encontradoProyecto = True
            Exit Do
        End If
        filaOT_Proyecto = filaOT_Proyecto + 4
        If filaOT_Proyecto > 10000 Then Exit Do
    Loop
    
    ' Verificar si se encontró en PROYECTO
    If Not encontradoProyecto Then
        MsgBox "No se encontró la OT " & OT_Buscar & " en la hoja PROYECTO.", vbExclamation
        Exit Sub
    End If
    
    ' ========== VERIFICAR ESTADO DE ENSAYOS ==========
    If UCase(Trim(estadoVerificacionAplicaciones)) <> "VERIFICADO" Then
        MsgBox "La OT " & OT_Buscar & " no tiene los ensayos verificados." & vbCrLf & _
               "Estado actual en columna FP: " & estadoVerificacionAplicaciones & vbCrLf & _
               "No se puede crear el paquete de ensayos.", vbExclamation
        Exit Sub
    End If
    
    ' ========== AGREGAR HOJAS SEGÚN LAS VARIABLES AP1-AP8 ==========
    ' Agregar hoja AP1 si tiene valor
    If Trim(AP1) <> "" Then
        If SheetExists(AP1) Then
            hojascopiar.Add wbOrigen.Worksheets(AP1)
        Else
            MsgBox "Advertencia: La hoja '" & AP1 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja AP2 si tiene valor
    If Trim(AP2) <> "" Then
        If SheetExists(AP2) Then
            hojascopiar.Add wbOrigen.Worksheets(AP2)
        Else
            MsgBox "Advertencia: La hoja '" & AP2 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja AP3 si tiene valor
    If Trim(AP3) <> "" Then
        If SheetExists(AP3) Then
            hojascopiar.Add wbOrigen.Worksheets(AP3)
        Else
            MsgBox "Advertencia: La hoja '" & AP3 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja AP4 si tiene valor
    If Trim(AP4) <> "" Then
        If SheetExists(AP4) Then
            hojascopiar.Add wbOrigen.Worksheets(AP4)
        Else
            MsgBox "Advertencia: La hoja '" & AP4 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja AP5 si tiene valor
    If Trim(AP5) <> "" Then
        If SheetExists(AP5) Then
            hojascopiar.Add wbOrigen.Worksheets(AP5)
        Else
            MsgBox "Advertencia: La hoja '" & AP5 & "' no existe en el libro.", vbExclamation
        End If
    End If
    ' Agregar hoja AP6 si tiene valor
    If Trim(AP6) <> "" Then
        If SheetExists(AP6) Then
            hojascopiar.Add wbOrigen.Worksheets(AP6)
        Else
            MsgBox "Advertencia: La hoja '" & AP6 & "' no existe en el libro.", vbExclamation
        End If
    End If
    ' Agregar hoja AP7 si tiene valor
    If Trim(AP7) <> "" Then
        If SheetExists(AP7) Then
            hojascopiar.Add wbOrigen.Worksheets(AP7)
        Else
            MsgBox "Advertencia: La hoja '" & AP7 & "' no existe en el libro.", vbExclamation
        End If
    End If
    ' Agregar hoja AP8 si tiene valor
    If Trim(AP8) <> "" Then
        If SheetExists(AP8) Then
            hojascopiar.Add wbOrigen.Worksheets(AP8)
        Else
            MsgBox "Advertencia: La hoja '" & AP8 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Verificar si hay hojas para copiar
    If hojascopiar.Count = 0 Then
        MsgBox "No hay hojas de ensayos para copiar." & vbCrLf & _
               "Verifique que las variables T1F a T5F contengan nombres de hojas válidos.", vbExclamation
        Exit Sub
    End If
    
    ' ========== CREAR ESTRUCTURA DE CARPETAS ==========
    ' Definir ruta base
    rutaBase = "C:\Users\sleon\OneDrive - industriascts.com\Pruebas FAT\"
    
    ' Ruta de la carpeta del proyecto
    rutaCarpetaProyecto = rutaBase & nombreCarpetaProyecto & "\"
    If Not fso.FolderExists(rutaCarpetaProyecto) Then
        fso.CreateFolder rutaCarpetaProyecto
    End If
    
    ' Crear nombre de carpeta principal (OT): OT-Producto-Aplicacion-Fecha
    nombreCarpetaPrincipal = "OT-" & OT_Buscar & "-" & Producto & "-" & Aplicacion & "-" & fechaCreacion
    rutaCarpetaPrincipal = rutaCarpetaProyecto & nombreCarpetaPrincipal & "\"
    
    ' Crear carpeta de OT si no existe
    If Not fso.FolderExists(rutaCarpetaPrincipal) Then
        fso.CreateFolder rutaCarpetaPrincipal
    End If
    
    ' Crear carpeta "02.Ensayos"
    nombreCarpetaAplicaciones= "03.Aplicaciones"
    rutaCarpetaAplicaciones = rutaCarpetaPrincipal & nombreCarpetaAplicaciones & "\"
    
    If Not fso.FolderExists(rutaCarpetaAplicaciones) Then
        fso.CreateFolder rutaCarpetaAplicaciones
    End If
    
    ' ========== CREAR Y GUARDAR LIBRO DE EXCEL ==========
    ' Crear nuevo libro de Excel
    Set wbNuevo = Workbooks.Add
    
    ' Eliminar hojas predeterminadas del nuevo libro (excepto una)
    Application.DisplayAlerts = False
    Do While wbNuevo.Worksheets.Count > 1
        wbNuevo.Worksheets(wbNuevo.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Copiar las hojas seleccionadas al nuevo libro
    For Each hoja In hojascopiar
        hoja.Copy After:=wbNuevo.Worksheets(wbNuevo.Worksheets.Count)
    Next hoja
    
    ' Eliminar la hoja vacía inicial
    Application.DisplayAlerts = False
    If wbNuevo.Worksheets.Count > 1 Then
        wbNuevo.Worksheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Preparar nombre de archivo
    nombreArchivo = OT_Buscar & "-" & "Paquete aplicacion.xlsx"
    
    ' Limpiar caracteres inválidos del nombre de archivo
    nombreArchivo = Replace(nombreArchivo, "/", "-")
    nombreArchivo = Replace(nombreArchivo, "\", "-")
    nombreArchivo = Replace(nombreArchivo, ":", "-")
    nombreArchivo = Replace(nombreArchivo, "*", "-")
    nombreArchivo = Replace(nombreArchivo, "?", "-")
    nombreArchivo = Replace(nombreArchivo, """", "-")
    nombreArchivo = Replace(nombreArchivo, "<", "-")
    nombreArchivo = Replace(nombreArchivo, ">", "-")
    nombreArchivo = Replace(nombreArchivo, "|", "-")
    
    ' Ruta completa del archivo final
    rutaCompleta = rutaCarpetaAplicaciones & nombreArchivo
    
    ' Ruta temporal para guardar primero
    rutaTemporal = Environ("TEMP") & "\" & nombreArchivo
    
    ' Eliminar archivo temporal si ya existe
    On Error Resume Next
    If Dir(rutaTemporal) <> "" Then Kill rutaTemporal
    On Error GoTo 0
    
    ' Guardar en ubicación temporal
    On Error GoTo ErrorGuardar
    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Eliminar archivo en destino final si existe
    On Error Resume Next
    If Dir(rutaCompleta) <> "" Then Kill rutaCompleta
    On Error GoTo 0
    
    ' Mover el archivo de temporal a ubicación final
    On Error GoTo ErrorMover
    fso.MoveFile rutaTemporal, rutaCompleta
    On Error GoTo 0
    
    ' ========== MENSAJE FINAL ==========
    MsgBox "Paquete de ensayos creado exitosamente en:" & vbCrLf & _
           rutaCarpetaAplicaciones & nombreArchivo & vbCrLf & vbCrLf & _
           "Proyecto: " & nombreCarpetaProyecto & vbCrLf & _
           "Hojas copiadas: " & hojascopiar.Count, vbInformation, "Proceso completado"
    
    ' Abrir la carpeta
    Shell "explorer.exe """ & rutaCarpetaAplicaciones & """", vbNormalFocus
    
    Exit Sub

    ErrorGuardar:
        Application.DisplayAlerts = True
        MsgBox "Error al guardar el archivo temporal:" & vbCrLf & _
            "Ruta: " & rutaTemporal & vbCrLf & vbCrLf & _
            "Error: " & Err.Description, vbCritical
        
        On Error Resume Next
        wbNuevo.Close SaveChanges:=False
        Exit Sub

    ErrorMover:
        MsgBox "Error al mover el archivo a la ubicación final:" & vbCrLf & _
            "Desde: " & rutaTemporal & vbCrLf & _
            "Hacia: " & rutaCompleta & vbCrLf & vbCrLf & _
            "Error: " & Err.Description & vbCrLf & vbCrLf & _
            "El archivo se guardó en la carpeta temporal." & vbCrLf & _
            "Puede moverlo manualmente.", vbCritical
        
        ' Abrir carpeta temporal
        Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Exit Sub
End Sub

Public Sub crearPaqueteEquipos()
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim hojascopiar As Collection
    Dim hoja As Worksheet
    Dim rutaBase As String
    Dim nombreCarpetaProyecto As String
    Dim nombreCarpetaPrincipal As String
    Dim nombreCarpetaequipos As String
    Dim rutaCarpetaProyecto As String
    Dim rutaCarpetaPrincipal As String
    Dim rutaCarpetaEquipos As String
    Dim nombreArchivo As String
    Dim estadoVerificacionEquipos As String
    Dim hoja_Proyecto As Worksheet
    Dim filaOT_Proyecto As Long
    Dim fechaCreacion As String
    Dim fso As Object
    Dim encontradoProyecto As Boolean
    Dim rutaTemporal As String
    Dim rutaCompleta As String
    
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection
    Set hoja_Proyecto = wbOrigen.Worksheets("PROYECTO")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la fecha actual en formato YYYY-MM-DD
    fechaCreacion = Format(Date, "yyyy-mm-dd")
    
    ' ========== OBTENER NOMBRE DEL PROYECTO DESDE CT59 ==========
    nombreCarpetaProyecto = Trim(wbOrigen.Worksheets("PROYECTO").Range("CT59").Value)
    
    ' Validar que el nombre del proyecto no esté vacío
    If nombreCarpetaProyecto = "" Then
        MsgBox "La celda CT59 no contiene un nombre de proyecto." & vbCrLf & _
               "Por favor, ingrese un nombre de proyecto antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Limpiar caracteres inválidos del nombre de carpeta proyecto
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "/", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "\", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ":", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "*", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "?", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, """", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "<", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, ">", "-")
    nombreCarpetaProyecto = Replace(nombreCarpetaProyecto, "|", "-")
    
    ' ========== BUSCAR OT EN HOJA PROYECTO ==========
    encontradoProyecto = False
    filaOT_Proyecto = 96
    Do While hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value <> ""
        If hoja_Proyecto.Range("BK" & filaOT_Proyecto).Value = OT_Buscar Then
            ' Obtener el estado de verificación de equipos de la columna FP
            estadoVerificacionEquipos = hoja_Proyecto.Range("HQ" & filaOT_Proyecto).Value
            encontradoProyecto = True
            Exit Do
        End If
        filaOT_Proyecto = filaOT_Proyecto + 4
        If filaOT_Proyecto > 10000 Then Exit Do
    Loop
    
    ' Verificar si se encontró en PROYECTO
    If Not encontradoProyecto Then
        MsgBox "No se encontró la OT " & OT_Buscar & " en la hoja PROYECTO.", vbExclamation
        Exit Sub
    End If
    
    ' ========== VERIFICAR ESTADO DE equipos ==========
    If UCase(Trim(estadoVerificacionEquipos)) <> "VERIFICADO" Then
        MsgBox "La OT " & OT_Buscar & " no tiene los ensayos verificados." & vbCrLf & _
               "Estado actual en columna Equipos: " & estadoVerificacionEquipos & vbCrLf & _
               "No se puede crear el paquete de ensayos.", vbExclamation
        Exit Sub
    End If
    ' ========== AGREGAR HOJAS SEGÚN LAS VARIABLES AP1-AP8 ==========
    ' Agregar hoja EQ1 si tiene valor
    If Trim(EQ1) <> "" Then
        If SheetExists(EQ1) Then
            hojascopiar.Add wbOrigen.Worksheets(EQ1)
        Else
            MsgBox "Advertencia: La hoja '" & EQ1 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja EQ2 si tiene valor
    If Trim(EQ2) <> "" Then
        If SheetExists(EQ2) Then
            hojascopiar.Add wbOrigen.Worksheets(EQ2)
        Else
            MsgBox "Advertencia: La hoja '" & EQ2 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja AP3 si tiene valor
    If Trim(EQ3) <> "" Then
        If SheetExists(EQ3) Then
            hojascopiar.Add wbOrigen.Worksheets(EQ3)
        Else
            MsgBox "Advertencia: La hoja '" & EQ3 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Agregar hoja EQ4 si tiene valor
    If Trim(EQ4) <> "" Then
        If SheetExists(EQ4) Then
            hojascopiar.Add wbOrigen.Worksheets(EQ4)
        Else
            MsgBox "Advertencia: La hoja '" & AP4 & "' no existe en el libro.", vbExclamation
        End If
    End If
    
    ' Verificar si hay hojas para copiar
    If hojascopiar.Count = 0 Then
        MsgBox "No hay hojas de formatos de equi para copiar." & vbCrLf & _
               "Verifique que las variables T1F a T5F contengan nombres de hojas válidos.", vbExclamation
        Exit Sub
    End If
    
    ' ========== CREAR ESTRUCTURA DE CARPETAS ==========
    ' Definir ruta base
    rutaBase = "C:\Users\sleon\OneDrive - industriascts.com\Pruebas FAT\"
    
    ' Ruta de la carpeta del proyecto
    rutaCarpetaProyecto = rutaBase & nombreCarpetaProyecto & "\"
    If Not fso.FolderExists(rutaCarpetaProyecto) Then
        fso.CreateFolder rutaCarpetaProyecto
    End If
    
    ' Crear nombre de carpeta principal (OT): OT-Producto-Aplicacion-Fecha
    nombreCarpetaPrincipal = "OT-" & OT_Buscar & "-" & Producto & "-" & Aplicacion & "-" & fechaCreacion
    rutaCarpetaPrincipal = rutaCarpetaProyecto & nombreCarpetaPrincipal & "\"
    
    ' Crear carpeta de OT si no existe
    If Not fso.FolderExists(rutaCarpetaPrincipal) Then
        fso.CreateFolder rutaCarpetaPrincipal
    End If
    
    ' Crear carpeta "02.Ensayos"
    nombreCarpetaequipos = "03.Aplicaciones"
    rutaCarpetaEquipos = rutaCarpetaPrincipal & nombreCarpetaequipos & "\"
    
    If Not fso.FolderExists(rutaCarpetaEquipos) Then
        fso.CreateFolder rutaCarpetaEquipos
    End If
    
    ' ========== CREAR Y GUARDAR LIBRO DE EXCEL ==========
    ' Crear nuevo libro de Excel
    Set wbNuevo = Workbooks.Add
    
    ' Eliminar hojas predeterminadas del nuevo libro (excepto una)
    Application.DisplayAlerts = False
    Do While wbNuevo.Worksheets.Count > 1
        wbNuevo.Worksheets(wbNuevo.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Copiar las hojas seleccionadas al nuevo libro
    For Each hoja In hojascopiar
        hoja.Copy After:=wbNuevo.Worksheets(wbNuevo.Worksheets.Count)
    Next hoja
    
    ' Eliminar la hoja vacía inicial
    Application.DisplayAlerts = False
    If wbNuevo.Worksheets.Count > 1 Then
        wbNuevo.Worksheets(1).Delete
    End If
    Application.DisplayAlerts = True
    
    ' Preparar nombre de archivo
    nombreArchivo = OT_Buscar & "-" & "Paquete aplicacion.xlsx"
    
    ' Limpiar caracteres inválidos del nombre de archivo
    nombreArchivo = Replace(nombreArchivo, "/", "-")
    nombreArchivo = Replace(nombreArchivo, "\", "-")
    nombreArchivo = Replace(nombreArchivo, ":", "-")
    nombreArchivo = Replace(nombreArchivo, "*", "-")
    nombreArchivo = Replace(nombreArchivo, "?", "-")
    nombreArchivo = Replace(nombreArchivo, """", "-")
    nombreArchivo = Replace(nombreArchivo, "<", "-")
    nombreArchivo = Replace(nombreArchivo, ">", "-")
    nombreArchivo = Replace(nombreArchivo, "|", "-")
    
    ' Ruta completa del archivo final
    rutaCompleta = rutaCarpetaAplicaciones & nombreArchivo
    
    ' Ruta temporal para guardar primero
    rutaTemporal = Environ("TEMP") & "\" & nombreArchivo
    
    ' Eliminar archivo temporal si ya existe
    On Error Resume Next
    If Dir(rutaTemporal) <> "" Then Kill rutaTemporal
    On Error GoTo 0
    
    ' Guardar en ubicación temporal
    On Error GoTo ErrorGuardar
    Application.DisplayAlerts = False
    wbNuevo.SaveAs Filename:=rutaTemporal, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNuevo.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Eliminar archivo en destino final si existe
    On Error Resume Next
    If Dir(rutaCompleta) <> "" Then Kill rutaCompleta
    On Error GoTo 0
    
    ' Mover el archivo de temporal a ubicación final
    On Error GoTo ErrorMover
    fso.MoveFile rutaTemporal, rutaCompleta
    On Error GoTo 0
    
    ' ========== MENSAJE FINAL ==========
    MsgBox "Paquete de ensayos creado exitosamente en:" & vbCrLf & _
           rutaCarpetaEquipos & nombreArchivo & vbCrLf & vbCrLf & _
           "Proyecto: " & nombreCarpetaProyecto & vbCrLf & _
           "Hojas copiadas: " & hojascopiar.Count, vbInformation, "Proceso completado"
    
    ' Abrir la carpeta
    Shell "explorer.exe """ & rutaCarpetaEquipos & """", vbNormalFocus
    
    Exit Sub

    ErrorGuardar:
        Application.DisplayAlerts = True
        MsgBox "Error al guardar el archivo temporal:" & vbCrLf & _
            "Ruta: " & rutaTemporal & vbCrLf & vbCrLf & _
            "Error: " & Err.Description, vbCritical
        
        On Error Resume Next
        wbNuevo.Close SaveChanges:=False
        Exit Sub

    ErrorMover:
        MsgBox "Error al mover el archivo a la ubicación final:" & vbCrLf & _
            "Desde: " & rutaTemporal & vbCrLf & _
            "Hacia: " & rutaCompleta & vbCrLf & vbCrLf & _
            "Error: " & Err.Description & vbCrLf & vbCrLf & _
            "El archivo se guardó en la carpeta temporal." & vbCrLf & _
            "Puede moverlo manualmente.", vbCritical
        
        ' Abrir carpeta temporal
        Shell "explorer.exe """ & Environ("TEMP") & """", vbNormalFocus
        Exit Sub
End Sub
' Función auxiliar para verificar si existe la hoja
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

PUBLIC SUB EJECUTOR()

    CALL Tomar_valores
    call OpcionInputBoxConValidacion
End SUB

Sub OpcionInputBoxConValidacion()
    Dim respuesta As String
    Dim mensaje As String
    Dim nombreCarpeta As String
    Dim valido As Boolean
    
    valido = False
    
    Do While Not valido
        mensaje = "CREAR CARPETA EN OT: " & OT_Buscar & vbCrLf & _
                  String(50, "=") & vbCrLf & vbCrLf & _
                  "Opciones disponibles:" & vbCrLf & vbCrLf & _
                  "  [01] Aseguramiento" & vbCrLf & _
                  "  [02] Ensayos" & vbCrLf & _
                  "  [03] Aplicaciones" & vbCrLf & _
                  "  [04] Equipos" & vbCrLf & vbCrLf & _
                  "Ingrese su opción:"
        
        respuesta = InputBox(mensaje, "Seleccionar Carpeta")
        
        ' Si cancela, salir
        If respuesta = "" Then
            MsgBox "Operación cancelada.", vbInformation
            Exit Sub
        End If
        
        ' Validar respuesta
        Select Case Trim(respuesta)
            Case "01", "1"
                nombreCarpeta = "01.Aseguramiento"
                valido = True
                CALL crearPaqueteAseguramiento
            Case "02", "2"
                nombreCarpeta = "02.Ensayos"
                valido = True
                CALL crearPaqueteEnsayos
            Case "03", "3"
                nombreCarpeta = "03.Aplicaciones"
                valido = True
                CALL crearPaqueteAplicaciones
            Case "04", "4"
                nombreCarpeta = "04.Equipos"
                valido = True
                CALL crearPaqueteEquipos
            Case Else
                MsgBox "✘ Opción no válida." & vbCrLf & vbCrLf & _
                       "Por favor ingrese un número entre 01 y 04.", _
                       vbExclamation, "Error de validación"
                ' El bucle continuará
        End Select
    Loop
    
    MsgBox "✓ Carpeta seleccionada: " & nombreCarpeta, vbInformation
End Sub

' RESTABLECER TABLA DE PROYECTO
PUBLIC SUB RESTABLECER_TABLA_PROYECTO()
    ThisWorkbook.Worksheets("PROYECTO").Range("CT59").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("GC64").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("BK96:BK156").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("CE96:CE156").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("DL96:DL156").Value = ""
END SUB
' RESTABLECER TABLA INSPECCION
Public Sub RESTABLECER_TABLA_INSPECCION()
    Dim i As Long
    Dim filaActual As Long
    Dim columnas As Variant
    
    With ThisWorkbook.Worksheets("INSPECCION")
        ' Limpiar rango de datos
        .Range("BA84:DB140").Value = ""
        
        ' Valores iniciales en fila 84
        .Range("EF84").Value = "FR-XXXX"
        .Range("ER84").Value = "FR-XXXX"
        .Range("FD84").Value = "FR-XXXX"
        .Range("GB84").Value = "FR-XXXX"
        .Range("GZ84").Value = "FR-XXXX"
        .Range("HX84").Value = "FR-XXXX"
        .Range("EO84:EO140").value = false 
        .Range("FA84:FA140").value = false
        .Range("FY84:FY140").value = false
        .Range("GW84:GW140").value = false
        .Range("HU84:HU140").value = false
        .Range("IS84:IS140").value = false
        
        ' Definir columnas a actualizar
        columnas = Array("EF", "ER", "FD", "GB", "GZ", "HX")
        
        ' Bucle para filas desde 88 hasta 140, de 4 en 4
        For i = 88 To 140 Step 4
            filaActual = i - 4  ' Fila anterior (referencia en la fórmula)
            
            ' Aplicar fórmula a cada columna
            Dim col As Variant
            For Each col In columnas
                .Range(col & i).FormulaLocal = "=SI(BA" & filaActual & "<>"""";""FR-XXXX"";"""")"
            Next col
        Next i
    End With
End Sub

sub limpiar_menu_Principal()
    ThisWorkbook.Worksheets("PROYECTO").Range("CT59").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("GC64").Value = ""
end sub




