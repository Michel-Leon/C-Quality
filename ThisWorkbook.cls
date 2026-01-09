'=========== Hoja principal del libro ============
' ============================================================================
' VARIABLES GLOBALES
' ============================================================================
Dim hojaAnterior As String

' ============================================================================
' EVENTO: ABRIR LIBRO
' ============================================================================
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' --- Paso 1: Ocultar todas las hojas excepto INICIO ---
    OcultarTodasLasHojas
    
    ' --- Paso 2: Mostrar y activar hoja INICIO ---
    MostrarHojaInicio
    
    ' --- Paso 3: Configurar interfaz ---
    ConfigurarInterfaz
    
    ' --- Paso 4: Cargar valores iniciales ---
    'CargarValoresIniciales
    
    ' --- Paso 5: Mostrar formulario de inicio ---
    MostrarFormularioInicio
    
End Sub

' ============================================================================
' EVENTO: SEGUIR HIPERVÍNCULO
' ============================================================================
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim nombreHoja As String
    
    On Error GoTo ManejadorError
    
    ' Obtener nombre de la hoja destino
    nombreHoja = Replace(Split(Target.SubAddress, "!")(0), "'", "")
    
    ' Validar que la hoja existe
    If Not SheetExists(nombreHoja) Then
        MsgBox "La hoja '" & nombreHoja & "' no existe.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Mostrar hoja destino
    With Sheets(nombreHoja)
        .Visible = xlSheetVisible
        .Activate
    End With
    
    ' Ocultar hoja anterior
    If hojaAnterior <> nombreHoja And hojaAnterior <> "" Then
        On Error Resume Next
        Sheets(hojaAnterior).Visible = xlSheetVeryHidden
        On Error GoTo 0
    End If
    
    ' Actualizar hoja anterior
    hojaAnterior = nombreHoja
    
    ManejadorError:
        If Err.Number <> 0 Then
            MsgBox "Error al navegar: " & Err.Description, vbCritical, "Error"
        End If
End Sub

' ============================================================================
' EVENTO: ANTES DE CERRAR
' ============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' Restaurar interfaz de Excel
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    
    With ActiveWindow
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
    On Error GoTo 0
End Sub

' ============================================================================
' PROCEDIMIENTOS PRIVADOS
' ============================================================================

' Oculta todas las hojas excepto INICIO
Private Sub OcultarTodasLasHojas()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "INICIO" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

' Muestra y activa la hoja INICIO
Private Sub MostrarHojaInicio()
    With Sheets("INICIO")
        .Visible = xlSheetVisible
        .Activate
    End With
    
    hojaAnterior = "INICIO"
End Sub

' Configura la interfaz ocultando elementos de Excel
Private Sub ConfigurarInterfaz()
    On Error Resume Next
    
    ' Ocultar Ribbon y barras
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    
    ' Ocultar elementos de ventana
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    
    On Error GoTo 0
End Sub

' Carga valores iniciales de configuración
'Private Sub CargarValoresIniciales()
    'On Error Resume Next
    
    ' Cargar valor del ComboBox de color si las hojas existen
    'If SheetExists("Vista frontal") And SheetExists("PORTADA") Then
        'Sheets("Vista frontal").Colorselec.Value = Sheets("PORTADA").Range("EU84").Value
    'End If
    
    'On Error GoTo 0
'End Sub

' Muestra el formulario de inicio
Private Sub MostrarFormularioInicio()
    On Error Resume Next
    INICIO_APP.Show
    On Error GoTo 0
End Sub

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

' Verifica si existe una hoja con el nombre dado
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
