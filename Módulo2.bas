Attribute VB_Name = "M�dulo2"
Sub OcultarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveSheet.Shapes("Mostrar Barras").Visible = True
    ActiveSheet.Shapes("ocultar Barras").Visible = False
End Sub
' Mostrar las barras y elementos de la interfaz
Sub MostarBARRAS()
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", true)"
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.DisplayHeadings = true
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveSheet.Shapes("ocultar Barras").Visible = True
    ActiveSheet.Shapes("Mostrar Barras").Visible = False
End Sub
Sub CambiarZoomEnTodasLasHojas()
    Dim zoomStr As String
    Dim zoomVal As Integer
    Dim ws As Worksheet

    ' Solicitar el valor de zoom al usuario
    zoomStr = InputBox("Ingrese el nivel de zoom deseado (ej. 100 para 100%):", "Cambiar Zoom")

    ' Cancelar si el usuario presiona Cancelar o deja vacío
    If zoomStr = "" Then Exit Sub

    ' Validar que el valor ingresado sea numérico
    If Not IsNumeric(zoomStr) Then
        MsgBox "Por favor ingrese un número válido.", vbExclamation
        Exit Sub
    End If

    zoomVal = CInt(zoomStr)

    ' Validar que el zoom esté dentro de un rango razonable
    If zoomVal < 10 Or zoomVal > 400 Then
        MsgBox "El nivel de zoom debe estar entre 10 y 400.", vbExclamation
        Exit Sub
    End If

    ' Aplicar el zoom a todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.Zoom = zoomVal
    Next ws

    MsgBox "Zoom ajustado a " & zoomVal & "% en todas las hojas.", vbInformation
End Sub

Sub MostrarTodasLasHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", TRUE)"
    Application.DisplayFormulaBar = True
End Sub

public sub Ensayos_Basicos()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Ensayos Basicos"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Ensayos Estandar").Visible = False
        activeSheet.Shapes("Ensayos PRO").Visible = False
        activeSheet.Shapes("Ensayos ENTERPRISE").Visible = False
        activeSheet.Shapes("Ensayo Personalizados").Visible = False
    End If
End Sub
public sub Ensayos_Estandar()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Ensayos Estandar"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Ensayos Basicos").Visible = False
        activeSheet.Shapes("Ensayos PRO").Visible = False
        activeSheet.Shapes("Ensayos ENTERPRISE").Visible = False
        activeSheet.Shapes("Ensayo Personalizados").Visible = False
    End If
End Sub
public sub Ensayos_PRO()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Ensayos PRO"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Ensayos Basicos").Visible = False
        activeSheet.Shapes("Ensayos Estandar").Visible = False
        activeSheet.Shapes("Ensayos ENTERPRISE").Visible = False
        activeSheet.Shapes("Ensayo Personalizados").Visible = False
    End If
End Sub 
public sub Ensayos_ENTERPRISE()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Ensayos ENTERPRISE"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Ensayos Basicos").Visible = False
        activeSheet.Shapes("Ensayos Estandar").Visible = False
        activeSheet.Shapes("Ensayos PRO").Visible = False
        activeSheet.Shapes("Ensayo Personalizados").Visible = False
    End If
End Sub
public sub Ensayos_Personalizados()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Ensayo Personalizados"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Ensayos Basicos").Visible = False
        activeSheet.Shapes("Ensayos Estandar").Visible = False
        activeSheet.Shapes("Ensayos PRO").Visible = False
        activeSheet.Shapes("Ensayos ENTERPRISE").Visible = False
    End If
End Sub
'----------------- MENUS PARA APLICAICONES -----------------
public sub I_MCC()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo I-MCC"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo L-Meter").Visible = False
        activeSheet.Shapes("Grupo O-Paralel").Visible = False
        activeSheet.Shapes("Grupo P-ATS").Visible = False
        activeSheet.Shapes("Grupo Q-Bank").Visible = False
        activeSheet.Shapes("Grupo V-Drive").Visible = False
    End If
End Sub
public sub L_METER()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo L-Meter"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo I-MCC").Visible = False
        activeSheet.Shapes("Grupo O-Paralel").Visible = False
        activeSheet.Shapes("Grupo P-ATS").Visible = False
        activeSheet.Shapes("Grupo Q-Bank").Visible = False
        activeSheet.Shapes("Grupo V-Drive").Visible = False
    End If
End Sub
public sub O_PARALEL()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo O-Paralel"

    on Error resume next    
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo I-MCC").Visible = False
        activeSheet.Shapes("Grupo L-Meter").Visible = False
        activeSheet.Shapes("Grupo P-ATS").Visible = False
        activeSheet.Shapes("Grupo Q-Bank").Visible = False
        activeSheet.Shapes("Grupo V-Drive").Visible = False
    End If
End Sub 
public sub P_ATS()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo P-ATS"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo I-MCC").Visible = False
        activeSheet.Shapes("Grupo L-Meter").Visible = False
        activeSheet.Shapes("Grupo O-Paralel").Visible = False
        activeSheet.Shapes("Grupo Q-Bank").Visible = False
        activeSheet.Shapes("Grupo V-Drive").Visible = False
    End If
End Sub
public sub Q_BANK()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo Q-Bank"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo I-MCC").Visible = False
        activeSheet.Shapes("Grupo L-Meter").Visible = False
        activeSheet.Shapes("Grupo O-Paralel").Visible = False
        activeSheet.Shapes("Grupo P-ATS").Visible = False
        activeSheet.Shapes("Grupo V-Drive").Visible = False
    End If
End Sub
public sub V_DRIVE()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Grupo V-Drive"

    on Error resume next    
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Grupo I-MCC").Visible = False
        activeSheet.Shapes("Grupo L-Meter").Visible = False
        activeSheet.Shapes("Grupo O-Paralel").Visible = False
        activeSheet.Shapes("Grupo P-ATS").Visible = False
        activeSheet.Shapes("Grupo Q-Bank").Visible = False
    End If
End Sub

'----------------- MENUS PARA SECCION DE EQUIPOS -----------------
public sub Sistema_de_proteccion()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Sistema de proteccion"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Conversion y distribucion de energia").Visible = False
        activeSheet.Shapes("Dispositivos de supervision y control").Visible = False
        activeSheet.Shapes("Instrumentacion y medicion").Visible = False
        activeSheet.Shapes("Servicios Auxiliares").Visible = False
    End If
End Sub
public sub Conversion_y_distribucion_de_energia()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Conversion y distribucion de energia"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Sistema de proteccion").Visible = False
        activeSheet.Shapes("Dispositivos de supervision y control").Visible = False
        activeSheet.Shapes("Instrumentacion y medicion").Visible = False
        activeSheet.Shapes("Servicios Auxiliares").Visible = False
    End If
End Sub
public sub Dispositivos_de_supervision_y_control()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Dispositivos de supervision y control"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Sistema de proteccion").Visible = False
        activeSheet.Shapes("Conversion y distribucion de energia").Visible = False
        activeSheet.Shapes("Instrumentacion y medicion").Visible = False
        activeSheet.Shapes("Servicios Auxiliares").Visible = False
    End If
End Sub
public sub Instrumentacion_y_medicion()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Instrumentacion y medicion"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Sistema de proteccion").Visible = False
        activeSheet.Shapes("Conversion y distribucion de energia").Visible = False
        activeSheet.Shapes("Dispositivos de supervision y control").Visible = False
        activeSheet.Shapes("Servicios Auxiliares").Visible = False
    End If
End Sub
public sub Servicios_Auxiliares()
    Dim shp As Shape
    Dim gruponame As String
    gruponame = "Servicios Auxiliares"

    on Error resume next
    set shp = ActiveSheet.Shapes(gruponame)
    On Error GoTo 0
    If not shp Is Nothing Then
        shp.Visible = Not shp.Visible
        activeSheet.Shapes("Sistema de proteccion").Visible = False
        activeSheet.Shapes("Conversion y distribucion de energia").Visible = False
        activeSheet.Shapes("Dispositivos de supervision y control").Visible = False
        activeSheet.Shapes("Instrumentacion y medicion").Visible = False
    End If
End Sub

' formularios de acceso

Sub AbrirLogin1()
    Bloqueo_Aseguramient.Show
End Sub

Sub AbrirLogin2()
    Bloqueo_USUARIOS.Show
End Sub

Sub AbrirLogin3()
    Bloqueo_Aplicaciones.Show
End Sub
sub AbrirLogin4()
    Bloqueo_Equipos.Show
End Sub
sub AbrirLogin5()
    Bloqueo_Ensayos.Show   
End Sub