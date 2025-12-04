Attribute VB_Name = "M�dulo1"
'=======VARIABLES GLOBALES======
public Cliente as string
public Elaborado_por as string
Public Producto As String
Public Aplicacion As String
PUBLIC OT_General As String
public OT_Buscar As String

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
            MsgBox "OT encontrada en la fila " & filaActual & vbCrLf & _
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
    Dim rutaCarpetaPrincipal As String
    Dim rutaCarpetaAseguramiento As String
    Dim nombreArchivo As String
    Dim estadoVerificacion As String
    Dim HOJA_INSPECCION As Worksheet
    Dim filaOT As Long
    Dim fechaCreacion As String
    Dim fso As Object
    
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection
    Set HOJA_INSPECCION = wbOrigen.Worksheets("INSPECCION")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la fecha actual en formato YYYY-MM-DD
    fechaCreacion = Format(Date, "yyyy-mm-dd")
    
    ' Buscar la fila de la OT para verificar el estado
    filaOT = 84
    Do While HOJA_INSPECCION.Range("BA" & filaOT).Value <> ""
        If HOJA_INSPECCION.Range("BA" & filaOT).Value = OT_Buscar Then
            estadoVerificacion = HOJA_INSPECCION.Range("EP" & filaOT).Value
            Exit Do
        End If
        filaOT = filaOT + 4
        If filaOT > 10000 Then Exit Do
    Loop
    
    ' Verificar si está "Verificado"
    If UCase(Trim(estadoVerificacion)) <> "VERIFICADO" Then
        MsgBox "La OT " & OT_Buscar & " no ha sido verificada." & vbCrLf & _
               "Estado actual: " & estadoVerificacion & vbCrLf & _
               "No se puede crear el paquete de aseguramiento.", vbExclamation
        Exit Sub
    End If
    
    ' Agregar hojas según los checkboxes
    Select Case CH1
        Case True, "Verdadero"
            Select Case columna1
                Case "FR-0856"
                    If SheetExists("FR-0856") Then hojascopiar.Add wbOrigen.Worksheets("FR-0856")
                Case "FR-0850"
                    If SheetExists("FR-0850") Then hojascopiar.Add wbOrigen.Worksheets("FR-0850")
            End Select
    End Select
    
    Select Case CH2
        Case True, "Verdadero"
            Select Case columna2
                Case "FR-0857"
                    If SheetExists("FR-0857") Then hojascopiar.Add wbOrigen.Worksheets("FR-0857")
                Case "FR-0851"
                    If SheetExists("FR-0851") Then hojascopiar.Add wbOrigen.Worksheets("FR-0851")
            End Select
    End Select
    
    Select Case CH3
        Case True, "Verdadero"
            Select Case columna3
                Case "FR-0858"
                    If SheetExists("FR-0858") Then hojascopiar.Add wbOrigen.Worksheets("FR-0858")
                Case "FR-0852"
                    If SheetExists("FR-0852") Then hojascopiar.Add wbOrigen.Worksheets("FR-0852")
            End Select
    End Select
    
    Select Case CH4
        Case True, "Verdadero"
            Select Case columna4
                Case "FR-0859"
                    If SheetExists("FR-0859") Then hojascopiar.Add wbOrigen.Worksheets("FR-0859")
                Case "FR-0853"
                    If SheetExists("FR-0853") Then hojascopiar.Add wbOrigen.Worksheets("FR-0853")
            End Select
    End Select
    
    Select Case CH5
        Case True, "Verdadero"
            Select Case columna5
                Case "FR-0860"
                    If SheetExists("FR-0860") Then hojascopiar.Add wbOrigen.Worksheets("FR-0860")
                Case "FR-0854"
                    If SheetExists("FR-0854") Then hojascopiar.Add wbOrigen.Worksheets("FR-0854")
            End Select
    End Select
    
    Select Case CH6
        Case True, "Verdadero"
            Select Case columna6
                Case "FR-0861"
                    If SheetExists("FR-0861") Then hojascopiar.Add wbOrigen.Worksheets("FR-0861")
                Case "FR-0855"
                    If SheetExists("FR-0855") Then hojascopiar.Add wbOrigen.Worksheets("FR-0855")
            End Select
    End Select
    
    ' Verificar si hay hojas para copiar
    If hojascopiar.Count = 0 Then
        MsgBox "No hay hojas seleccionadas para copiar.", vbExclamation
        Exit Sub
    End If
    
    ' Definir ruta base (CAMBIAR ESTA RUTA A LA UBICACIÓN DESEADA)
    rutaBase = "C:\Users\sleon\OneDrive - industriascts.com\Pruebas FAT"
    
    ' Crear nombre de carpeta principal: OT-Producto-Aplicacion-Fecha
    nombreCarpetaPrincipal = "OT-" & OT_Buscar & "-" & Producto & "-" & Aplicacion & "-" & fechaCreacion
    rutaCarpetaPrincipal = rutaBase & nombreCarpetaPrincipal & "\"
    
    ' Crear carpeta principal si no existe
    If Not fso.FolderExists(rutaCarpetaPrincipal) Then
        fso.CreateFolder rutaCarpetaPrincipal
    End If
    
    ' Crear carpeta "Aseguramiento"
    nombreCarpetaAseguramiento = "Aseguramiento"
    rutaCarpetaAseguramiento = rutaCarpetaPrincipal & nombreCarpetaAseguramiento & "\"
    
    If Not fso.FolderExists(rutaCarpetaAseguramiento) Then
        fso.CreateFolder rutaCarpetaAseguramiento
    End If
    
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
    wbNuevo.Worksheets(1).Delete
    Application.DisplayAlerts = True
    
    ' Guardar el nuevo libro
    nombreArchivo = OT_Buscar & "-"&"Paquete de Aseguramiento.xlsx"
    wbNuevo.SaveAs Filename:=rutaCarpetaAseguramiento & nombreArchivo, _
                   FileFormat:=xlOpenXMLWorkbook
    
    wbNuevo.Close SaveChanges:=False
    
    ' Mensaje de confirmación
    MsgBox "Paquete de aseguramiento creado exitosamente en:" & vbCrLf & _
           rutaCarpetaAseguramiento & nombreArchivo, vbInformation, "Proceso completado"
    
    ' Abrir la carpeta
    Shell "explorer.exe """ & rutaCarpetaAseguramiento & """", vbNormalFocus
    
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
    CALL crearPaqueteAseguramiento

    MsgBox "Proceso completado", vbInformation

End SUB












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



