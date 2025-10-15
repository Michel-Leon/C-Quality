'============inicio del programa inspeccion===================
Public productoSeleccionado As String

public Sub CrearPaquetePorProducto()

    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim hojascopiar AS collection
    Dim hoja AS worksheet
    Dim nombreHoja AS Variant
    Dim chk AS Object
    

    'asignar el libro actial como origen
    Set wbOrigen = ThisWorkbook
    Set hojascopiar = New Collection

    'Revisa los checkboxes y agrega las hojas correspondientes a la coleccion
    with wbOrigen.sheets("INSPECCION")

        select case Ucase(productoSeleccionado)
            case "T-POWER"
                hojasCopiar.Add "FR-0860"
                If MTM5 = True Then hojascopiar.Add "0860-MTM"
                If BRR5 = True Then hojascopiar.Add "0860-BRR"
                If CRP5 = True Then hojascopiar.Add "0860-CPT" 
                If CRT5 = True Then hojascopiar.Add "0860-CTR"
                If CMM5 = True Then hojascopiar.Add "0860-CMM"
            case "G-FLEX"
                hojasCopiar.Add "FR-0850"
                If MTM4 = True Then hojascopiar.Add "0860-MTM"
                If BRR4 = True Then hojascopiar.Add "0860-BRR"
                If CRP4 = True Then hojascopiar.Add "0860-CPT" 
                If CRT4 = True Then hojascopiar.Add "0860-CTR"
                If CMM4 = True Then hojascopiar.Add "0860-CMM"
            case "D-Critical"
                hojasCopiar.Add "FR-0819"
                If MTM2 = True Then hojascopiar.Add "0870-MTM"
                If BRR2 = True Then hojascopiar.Add "0870-BRR"
                If CRP2 = True Then hojascopiar.Add "0870-CPT" 
                If CRT2 = True Then hojascopiar.Add "0870-CTR"
                If CMM2 = True Then hojascopiar.Add "0870-CMM"
            case "M-Board"
                hojasCopiar.Add "FR-0816"
                If MTM1 = True Then hojascopiar.Add "0870-MTM"
                If BRR1 = True Then hojascopiar.Add "0870-BRR"
                If CRP1 = True Then hojascopiar.Add "0870-CPT" 
                If CRT1 = True Then hojascopiar.Add "0870-CTR"
                If CMM1 = True Then hojascopiar.Add "0870-CMM"
            case "BlockSet"
                hojasCopiar.Add "FR-08120"
                If MTM3 = True Then hojascopiar.Add "0870-MTM"
                If BRR3 = True Then hojascopiar.Add "0870-BRR"
                If CRP3 = True Then hojascopiar.Add "0870-CPT" 
                If CRT3 = True Then hojascopiar.Add "0870-CTR"
                If CMM3 = True Then hojascopiar.Add "0870-CMM"
            case Else
                'SI NO SE SELECCIONA  EL PAQUETE COMPLETO
                Dim crtl as OLEObject
                For Each ctrl In .OLEObjects
                    If TypeName(ctrl.Object) = "CheckBox" Then
                        If ctrl.Object.Value = True Then
                            Select Case UCase(ctrl.Name)
                                Case "MTM1", "MTM2", "MTM3", "MTM4", "MTM5"
                                    hojasCopiar.Add "0870-MTM"
                                Case "BRR1", "BRR2", "BRR3", "BRR4", "BRR5"
                                    hojasCopiar.Add "0870-BRR"
                                Case "CRP1", "CRP2", "CRP3", "CRP4", "CRP5"
                                    hojasCopiar.Add "0870-CPT"
                                Case "CRT1", "CRT2", "CRT3", "CRT4", "CRT5"
                                    hojasCopiar.Add "0870-CTR"
                                Case "CMM1", "CMM2", "CMM3", "CMM4", "CMM5"
                                    hojasCopiar.Add "0870-CMM"
                            End Select
                        End If
                    End If
                Next ctrl    
        end select
    End With
    'verificar sis e selecciono una hoja
    If hojascopiar.Count = 0 Then
        MsgBox "Seleccione al menos una hoja",vbExclamation, "Atencion"
        Exit Sub
    End If
    'crear nuevo libro
    Set wbNuevo = Workbooks.Add
    'copiar las hojas seleccionadas al nuevo libro
    For Each nombreHoja In hojascopiar
        on error resume next
        wbOrigen.Sheets(nombreHoja).Copy after:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
        If Err.Number <> 0 Then
            MsgBox "Error al copiar la hoja " & nombreHoja, vbExclamation, "Atencion"
            err.clear
        End If
        on error goto 0
    Next nombreHoja
    'eliminar hojas predeterminadas si existen
    Application.DisplayAlerts = False
    For Each hoja In wbNuevo.Sheets
        If hoja.Name Like "Hoja*" Or hoja.Name Like "Sheet*" Then hoja.Delete
    Next hoja
    Application.DisplayAlerts = True
    
    MsgBox  "Paquete creado correctamente para el producto: " & productoSeleccionado, vbInformation
End Sub
'========================================
Public Sub Seleccionar_Tpower()
    productoSeleccionado = "T-POWER"
    Dim estadoNuevo As Boolean
    
    ' Determinar el estado a aplicar:
    ' si al menos uno NO está marcado → marcar todos
    ' si todos están marcados → desmarcar todos
    If MTM5.Value = True And BRR5.Value = True And CRP5.Value = True And _
       CRT5.Value = True And CMM5.Value = True Then
        estadoNuevo = False
    Else
        estadoNuevo = True
    End If
    
    ' Aplicar el estado nuevo a todos los CheckBox
    MTM5.Value = estadoNuevo
    BRR5.Value = estadoNuevo
    CRP5.Value = estadoNuevo
    CRT5.Value = estadoNuevo
    CMM5.Value = estadoNuevo
End Sub
Public Sub Seleccionar_Gflex()
    productoSeleccionado = "G-FLEX"
    Dim estadoNuevo As Boolean
    
    ' Determinar el estado a aplicar:
    ' si al menos uno NO está marcado → marcar todos
    ' si todos están marcados → desmarcar todos
    If MTM4.Value = True And BRR4.Value = True And CRP4.Value = True And _
       CRT4.Value = True And CMM4.Value = True Then
        estadoNuevo = False
    Else
        estadoNuevo = True
    End If
    
    ' Aplicar el estado nuevo a todos los CheckBox
    MTM4.Value = estadoNuevo
    BRR4.Value = estadoNuevo
    CRP4.Value = estadoNuevo
    CRT4.Value = estadoNuevo
    CMM4.Value = estadoNuevo
End Sub
Public Sub Seleccionar_Dcritical()
    productoSeleccionado = "D-Critical"
    Dim estadoNuevo As Boolean
    
    ' Determinar el estado a aplicar:
    ' si al menos uno NO está marcado → marcar todos
    ' si todos está marcados → desmarcar todos
    If MTM2.Value = True And BRR2.Value = True And CRP2.Value = True And _
       CRT2.Value = True And CMM2.Value = True Then
        estadoNuevo = False
    Else
        estadoNuevo = True
    End If
    
    ' Aplicar el estado nuevo a todos los CheckBox
    MTM2.Value = estadoNuevo
    BRR2.Value = estadoNuevo
    CRP2.Value = estadoNuevo
    CRT2.Value = estadoNuevo
    CMM2.Value = estadoNuevo
End Sub
Public Sub Seleccionar_Mboard()
    productoSeleccionado = "M-Board"
    Dim estadoNuevo As Boolean
    
    ' Determinar el estado a aplicar:
    ' si al menos uno NO está marcado → marcar todos
    ' si todos están marcados → desmarcar todos
    If MTM1.Value = True And BRR1.Value = True And CRP1.Value = True And _
       CRT1.Value = True And CMM1.Value = True Then
        estadoNuevo = False
    Else
        estadoNuevo = True
    End If
    
    ' Aplicar el estado nuevo a todos los CheckBox
    MTM1.Value = estadoNuevo
    BRR1.Value = estadoNuevo
    CRP1.Value = estadoNuevo
    CRT1.Value = estadoNuevo
    CMM1.Value = estadoNuevo
End Sub
Public Sub Seleccionar_BlockSet()
    productoSeleccionado = "BlockSet"
    Dim estadoNuevo As Boolean
    
    ' Determinar el estado a aplicar:
    ' si al menos uno NO está marcado → marcar todos
    ' si todos está marcados → desmarcar todos
    If MTM3.Value = True And BRR3.Value = True And CRP3.Value = True And _
       CRT3.Value = True And CMM3.Value = True Then
        estadoNuevo = False
    Else
        estadoNuevo = True
    End If
    
    ' Aplicar el estado nuevo a todos los CheckBox
    MTM3.Value = estadoNuevo
    BRR3.Value = estadoNuevo
    CRP3.Value = estadoNuevo
    CRT3.Value = estadoNuevo
    CMM3.Value = estadoNuevo
End Sub
'========================================
Public SUB LIMPIARSELECCION()
    Dim ctrl As OLEObject
    Dim hoja As Worksheet
    
    ' 🔹 Cambia "INSPECCION" por el nombre de tu hoja
    Set hoja = ThisWorkbook.Sheets("INSPECCION")
    
    ' Recorre todos los controles ActiveX de la hoja
    For Each ctrl In hoja.OLEObjects
        If TypeName(ctrl.Object) = "CheckBox" Then
            ctrl.Object.Value = False
        End If
    Next ctrl
    
    MsgBox "Todos los CheckBox han sido desmarcados.", vbInformation
End Sub
public sub ejecutar_Tpower()
    Call CrearPaquetePorProducto
end sub
