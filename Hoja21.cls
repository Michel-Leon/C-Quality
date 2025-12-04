'=================HOJA PROYECTO==============================
public sub worksheet_activate()
    with Me.OLEObjects("Elaboracion").Object
        .clear
        .additem "Persona 1"
        .additem "Persona 2"
        .additem "Persona 3"
    end with
    with Me.OLEObjects("Responsable").Object 
        .clear
        .additem "Responsable 1"
        .additem "Responsable 2"
        .additem "Responsable 3"
    end with
    with Me.OLEObjects("Linea_Producto").Object
        .clear
        .additem "T-power"
        .additem "G-flex"
        .additem "BlokSeT"
    end with 
    with Me.OLEObjects("Aplicacion_Producto").Object
        .clear
        .additem "Distribución"
        .additem "Transferencia"
        .additem "Compensación"
        .additem "Sincronismo"
        .additem "Arranque Motor"
        .additem "PDU"
        .additem "RPP"
    end with
End Sub 

Public Sub Crear_Proyecto()
    Dim cliente As String
    Dim HOJA_PROYECTO As Worksheet
    Dim formatos_Destino As Variant
    Dim formato As Variant
    Dim filaDestino As Long
    
    Set HOJA_PROYECTO = ThisWorkbook.Worksheets("PROYECTO")
    Elaborado_por = Me.OLEObjects("Elaboracion").Object.Value
    Producto = Me.OLEObjects("Linea_Producto").Object.Value
    Aplicacion = Me.OLEObjects("Aplicacion_Producto").Object.Value

    With HOJA_PROYECTO
        cliente = .Range("CT59").Value
        OT_General = .Range("GC64").Value
        
        ' Buscar la primera fila vacía desde la fila 96, avanzando de 4 en 4
        filaDestino = 96
        Do While .Range("BK" & filaDestino).Value <> "" Or _
                 .Range("CE" & filaDestino).Value <> ""
            filaDestino = filaDestino + 4
        Loop

        ' Escribir en la fila encontrada
        .Range("BK" & filaDestino).Value = OT_General
        .Range("CE" & filaDestino).Value = Producto
    End With
    
    Select Case Producto
        Case "T-power"
            formatos_Destino = Array("FR-0856", "FR-0857", "FR-0858", "FR-0859", "FR-0860")
            'Recorre cada hoja del array
            For Each formato In formatos_Destino
                If SheetExists(CStr(formato)) Then
                    With ThisWorkbook.Worksheets(CStr(formato))
                        .Range("F13").Value = cliente
                        .Range("AQ13").Value = OT_General
                        .Range("H113").Value = Elaborado_por
                        .Range("Y13").Value = Producto & "-" & Aplicacion
                    End With
                End If
            Next formato
            
        Case "G-flex"
            formatos_Destino = Array("FR-0861", "FR-0862", "FR-0863", "FR-0864", "FR-0865")
            'Recorre cada hoja del array
            For Each formato In formatos_Destino
                If SheetExists(CStr(formato)) Then
                    With ThisWorkbook.Worksheets(CStr(formato))
                        .Range("F13").Value = cliente
                        .Range("AQ13").Value = OT_General
                        .Range("H113").Value = Elaborado_por
                        .Range("Y13").Value = Producto & "-" & Aplicacion
                    End With
                End If
            Next formato
    End Select
End Sub

' No olvides incluir la función SheetExists
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function