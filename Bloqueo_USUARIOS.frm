VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloqueo_USUARIOS 
   Caption         =   "LOGIN"
   ClientHeight    =   5120
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   4850
   OleObjectBlob   =   "Bloqueo_USUARIOS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bloqueo_USUARIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdLogin_Click()
   'Dim usuario As String, clave As String
    usuario = txtUsuario.Text
    clave = txtPassword.Text

    
    If usuario = "ADMIN" And clave = "Admin@GT78" Or usuario = "SLEON" And clave = "Tad@SL97" Then
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", TRUE)" 'BARRA DE INICIO
    Application.DisplayFormulaBar = True 'BARRA FORMULARIO
    ActiveWindow.DisplayHeadings = False ' ENCABEZADO
    ThisWorkbook.Names.Add Name:="usuario_actual", RefersTo:="=""" & usuario & """"
    Sheets("USUARIOS").Visible = True
    Sheets("USUARIOS").Activate
    Sheets("ACCESOS").Visible = True
    Sheets("ACCESOS").Activate
    Unload Me
    
Else
    MsgBox "Usuario o contraseña incorrectos", vbCritical
    txtPassword.Text = ""
    txtUsuario.SetFocus
End If

End Sub




Private Sub UserForm_Initialize()
    ' Centrado relativo en pantalla, o en un lugar espec?fico
    Me.StartUpPosition = 0 ' Manual
    Me.Top = 100   ' Distancia desde la parte superior de la pantalla (ajusta a gusto)
    Me.Left = 400  ' Distancia desde el borde izquierdo (ajusta a gusto)
End Sub



