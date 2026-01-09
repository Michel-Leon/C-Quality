VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bloqueo_Aseguramiento_V1 
   Caption         =   "LOGIN"
   ClientHeight    =   5130
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   4850
   OleObjectBlob   =   "Bloqueo_Aseguramiento_V1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Bloqueo_Aseguramiento_V1"
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
    Sheets("FR-0850").Visible = True
    Sheets("FR-0850").Activate
    Sheets("FR-0851").Visible = True
    Sheets("FR-0851").Activate 
    Sheets("FR-0852").Visible = True
    Sheets("FR-0852").Activate
    Sheets("FR-0853").Visible = True
    Sheets("FR-0853").Activate
    Sheets("FR-0854").Visible = True
    Sheets("FR-0854").Activate
    Sheets("FR-0856").Visible = True
    Sheets("FR-0856").Activate
    Sheets("FR-0857").Visible = True
    Sheets("FR-0857").Activate
    Sheets("FR-0858").Visible = True
    Sheets("FR-0858").Activate
    Sheets("FR-0859").Visible = True
    Sheets("FR-0859").Activate
    Sheets("FR-0860").Visible = True
    Sheets("FR-0860").Activate
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




