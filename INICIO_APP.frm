VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INICIO_APP 
   Caption         =   "LOGIN"
   ClientHeight    =   7610
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   6570
   OleObjectBlob   =   "INICIO_APP.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "INICIO_APP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdLogin_Click()
    Dim usuario As String, clave As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim loginCorrecto As Boolean
    Dim esSuperAdmin As Boolean
    Dim wsLog As Worksheet
    Dim filaLog As Long

    usuario = Trim(txtUsuario.Text)
    clave = Trim(txtPassword.Text)
    
    esSuperAdmin = False

    ' Super Admin codificado manualmente
    If usuario = "ADMIN" And clave = "Admin@GT78" Then
        esSuperAdmin = True
        loginCorrecto = True
    Else
        Set ws = ThisWorkbook.Sheets("Usuarios")
        lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        
        For i = 2 To lastRow
            If ws.Cells(i, 5).Value = usuario And ws.Cells(i, 6).Value = clave Then
                loginCorrecto = True
                Exit For
            End If
        Next i
    End If

    If loginCorrecto Then
        INICIO_APP.Hide
        Sheets("INICIO").Activate
        
        ' Registrar acceso
        Set wsLog = ThisWorkbook.Sheets("Accesos")
        filaLog = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
        wsLog.Cells(filaLog, 1).Value = Now                ' Fecha y hora
        wsLog.Cells(filaLog, 2).Value = usuario            ' Usuario
        wsLog.Cells(filaLog, 3).Value = Environ("Username") ' Usuario del sistema (opcional)

        ThisWorkbook.Save 'GUARDAE EL ARCHIVO AUTOMATICAMENTE

        If Not esSuperAdmin Then
            Sheets("INICIO").Range("HD4").Value = ws.Cells(i, 4).Value
        Else
            MsgBox "Acceso concedido como Super Administrador.", vbInformation
        End If
    Else
        MsgBox "El usuario o la clave son incorrectos.", vbExclamation
    End If
End Sub

