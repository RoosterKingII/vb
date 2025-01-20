VERSION 5.00
Begin VB.Form FrmAcceso 
   Caption         =   "Control de Acceso"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3825
   LinkTopic       =   "Form11"
   ScaleHeight     =   3405
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox CboPerfil 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Txtcontrasenausuario 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "X"
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Txtnombreusuario 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Perfil:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
    If funciones.cedusuario <> "" And funciones.codperfil <> 0 Then
        Txtcontrasenausuario.Text = ""
        Txtnombreusuario.Text = ""
        CboPerfil.Clear
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "Datos Incompletos o Erróneos"
        Txtcontrasenausuario.Text = ""
        Txtnombreusuario.Text = ""
        CboPerfil.Clear
    End If
End Sub

Private Sub CmdAceptar_KeyPress(keyascii As Integer)
    Call CmdAceptar_Click
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.SetFocus
  'ME.Height = 230
  'ME.Width = 380
    'funciones.CentrarenPantalla (Form11)
    CboPerfil.Clear
    'Txtnombreusuario.SetFocus
End Sub


Public Sub Txtcontrasenausuario_KeyPress(keyascii As Integer)

  Dim rs As Recordset
  Dim i As Integer
  Dim sql As String
  Dim cantperfiles As Integer
  
 ' On Error GoTo Error
  
  If keyascii = 13 Then
    If Txtnombreusuario.Text <> "" And Txtcontrasenausuario.Text <> "" Then 'si no son vacios estos objetos
      funciones.Conectar2          'se conecta a la base de datos de usuarios esparta
      sql = "select cedusuario as resultado from usuario where nicusuario='" & Txtnombreusuario.Text & "' and clausuario='" & LCase(Txtcontrasenausuario.Text) & "'"
      funciones.cedusuario = funciones.CampoString(sql, cn2)  'retorna la CI del usuario
      If funciones.cedusuario Then 'si encuentra a un usuario que corresponda
        sql = "SELECT count(perfil.desperfil) as resultado FROM perfilusuario, perfil, usuario WHERE perfilusuario.perfil_codperfil = perfil.codperfil AND" & _
              " usuario.cedusuario = perfilusuario.usuario_cedusuario AND usuario.cedusuario = '" & funciones.cedusuario & "'"
        cantperfiles = funciones.CampoEntero(sql, cn2)
        Select Case cantperfiles
          Case 0
            MsgBox "Usuario sin Perfil Definido", vbCritical
          Case 1
            sql = "SELECT perfil.desperfil as resultado FROM perfilusuario, perfil, usuario  WHERE perfilusuario.perfil_codperfil = perfil.codperfil AND" & _
              " usuario.cedusuario = perfilusuario.usuario_cedusuario AND usuario.cedusuario = '" & funciones.cedusuario & "'"
            Variables.desperfil = funciones.CampoString(sql, cn2)
            sql = "Select codperfil as resultado from perfil where desperfil='" & funciones.desperfil & "'"
            funciones.codperfil = funciones.CampoEntero(sql, cn2)
            BtnAceptar.Enabled = True
            BtnAceptar.SetFocus
          Case Else
            Label3.Visible = True
            CboPerfil.Visible = True
            sql = "SELECT perfil.desperfil as resultado FROM perfilusuario, perfil, usuario  WHERE perfilusuario.perfil_codperfil = perfil.codperfil AND" & _
              " usuario.cedusuario = perfilusuario.usuario_cedusuario AND usuario.cedusuario = '" & funciones.cedusuario & "'"
            funciones.llenarcombobox CboPerfil, sql, cn2, True
        End Select
      Else
        MsgBox "Usuario No Definido", vbCritical
        Txtcontrasenausuario.Text = ""
        Txtnombreusuario.Text = ""
        Txtnombreusuario.SetFocus
      End If
      cn2.Close 'cierra la bd
      
    End If
  End If
  Exit Sub
'Error:
'        MsgBox Err.Description, vbCritical
End Sub

Public Sub Txtnombreusuario_KeyPress(keyascii As Integer)

  If keyascii = 13 Then
    Txtcontrasenausuario.SetFocus
  End If
End Sub

Public Sub CboPerfil_Click()
  Dim RsPerfil As Recordset
  Dim sql As String
  
  If CboPerfil.Text <> "Seleccione" Then
    If Len(CboPerfil.Text) > 0 Then
      funciones.desperfil = CboPerfil.Text
      funciones.Conectar2
      sql = "Select codperfil as resultado from perfil where desperfil='" & funciones.desperfil & "'"
      funciones.codperfil = funciones.CampoEntero(sql, cn2)
      cn2.Close
      CmdAceptar.Enabled = True
      CmdAceptar.SetFocus
    End If
  Else
    CmdAceptar.Enabled = False
  End If
End Sub

Public Sub Txtnombreusuario_Change()
  Txtnombreusuario.Text = UCase(Txtnombreusuario.Text)
End Sub

