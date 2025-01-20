VERSION 5.00
Begin VB.Form FrmValidarFolio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validar Firma de Folio"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6855
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "&Procesar"
      Height          =   735
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox TxtLibroIni 
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtLibroFin 
         Height          =   375
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtHasta 
         Height          =   375
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtDesde 
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Libro fin"
         Height          =   195
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Folio fin"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio Inicio"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmValidarFolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdLimpiar_Click()
    TxtDesde.Text = ""
    TxtHasta.Text = ""
    TxtLibroIni.Text = ""
    TxtLibroFin.Text = ""
    TxtLibroIni.SetFocus
End Sub

Private Sub CmdProcesar_Click()
    Dim codparticipante As Long
    Dim sql As String
    Dim graduando As String
    Dim cedusuario As String
    Dim cedgraduando As String
    Dim i As Integer
    
    funciones.Conectar
    For i = Val(TxtDesde.Text) To Val(TxtHasta.Text)
        sql = "select participantes_codparticipantes as resultado from participantepromocion where liparticipantepromocion='" & TxtLibroIni.Text & "' and  foparticipantepromocion='" & i & "'" 'funciones.formatofolio(fol)
        codparticipante = funciones.CampoEnteroLargo(sql, cn)
        sql = "update participantepromocion set ffparticipantepromocion='true' where liparticipantepromocion='" & TxtLibroIni.Text & "' and  foparticipantepromocion='" & i & "'" '& " and participantes_codparticipantes=" & codparticipante
        cn.Execute (sql)
        sql = "select nomapeusuario as resultado from usuario, participante where usuario.cedusuario=participante.usuario_cedusuario and codparticipantes=" & codparticipante
        graduando = funciones.CampoString(sql, cn)
        sql = "select cedusuario as resultado from usuario, participante where usuario.cedusuario=participante.usuario_cedusuario and codparticipantes=" & codparticipante
        cedgraduando = funciones.CampoString(sql, cn)
        funciones.RegistroEvento funciones.cedusuario, Date, "Verificación de Folio Firmado " & funciones.formatofolio(i) & " del libro " & TxtLibroIni, funciones.usuario, cedgraduando, graduando
    Next
    cn.Close
    MsgBox "Proceso concluido", vbInformation
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
