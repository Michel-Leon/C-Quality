VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloqueo_Equipos 
   Caption         =   "LOGIN"
   ClientHeight    =   5130
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   4850
   OleObjectBlob   =   "Bloqueo_Equipos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bloqueo_Equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdLogin_Click()
   
    usuario = txtUsuario.Text
    clave = txtPassword.Text

    
    If usuario = "ADMIN" And clave = "Admin@GT78" Then
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ThisWorkbook.Names.Add Name:="usuario_actual", RefersTo:="=""" & usuario & """"
    Sheets("ACCESOSRIOS PUERTA").Visible = True
    Sheets("ACCESOSRIOS PUERTA").Activate
    Unload Me
    
Else
    MsgBox "Usuario o contraseña incorrectos", vbCritical
    txtPassword.Text = ""
    txtUsuario.SetFocus
End If

End Sub

Private Sub UserForm_Initialize()
   
    Me.StartUpPosition = 0
    Me.Top = 100
    Me.Left = 400
End Sub




