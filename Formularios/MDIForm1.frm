VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnusalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuactas 
      Caption         =   "Actas Continuas"
      Begin VB.Menu mnugeneraractas 
         Caption         =   "Regular"
      End
      Begin VB.Menu mnugeneraractasPER 
         Caption         =   "P.E.R"
      End
   End
   Begin VB.Menu mnuprocesos 
      Caption         =   "&Procesos"
      Begin VB.Menu Mnumigrar 
         Caption         =   "Notas Mision Sucre"
      End
   End
   Begin VB.Menu mnuconstancias 
      Caption         =   "&Constancias"
      Begin VB.Menu mnunotas 
         Caption         =   "Generar"
      End
   End
   Begin VB.Menu mnugrado 
      Caption         =   "&Grado"
      Begin VB.Menu mnufolios 
         Caption         =   "Generar Folios"
      End
      Begin VB.Menu mnulistado 
         Caption         =   "Listado Promoción"
      End
      Begin VB.Menu mnuValidar 
         Caption         =   "Validar Firma de Folio"
      End
      Begin VB.Menu mnuimprimirtitulos 
         Caption         =   "Impresión de Titulos"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Dim sql As String
  
  
  funciones.Conectar
  sql = "select nomusuario as resultado from usuario where cedusuario='" & funciones.cedusuario & "'"
  funciones.usuario = funciones.CampoString(sql, cn)
  sql = "select apeusuario as resultado from usuario where cedusuario='" & funciones.cedusuario & "'"
  funciones.usuario = funciones.usuario & " " & funciones.CampoString(sql, cn)
  MDIForm1.Caption = MDIForm1.Caption & " <<" & funciones.usuario & ", " & funciones.desperfil & ">>"
  MDIForm1.Visible = False
  'ConfigurarMenu()
  cn.Close
  '/* seguridad * /
  funciones.RegistroEvento funciones.cedusuario, funciones.FormatoFechaConsulta(Date), "Ingreso al Sistema", funciones.usuario, "", ""

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    '/* seguridad * /
    funciones.RegistroEvento funciones.cedusuario, funciones.FormatoFechaConsulta(Date), "Salida del Sistema", funciones.usuario, "", ""
End Sub

Private Sub mnufolios_Click()
    Load FrmGenerarFolios
    FrmGenerarFolios.Show
End Sub

Private Sub mnugeneraractas_Click()
    Load FrmActaContinua
    FrmActaContinua.Show
End Sub

Private Sub mnugeneraractasPER_Click()
    Load FrmActaContinua
    FrmActaContinua.Caption = FrmActaContinua.Caption & " P.E.R"
    FrmActaContinua.Show
End Sub

Private Sub mnuimprimirtitulos_Click()
    Load FrmImprimirTitulos
    FrmImprimirTitulos.Show
End Sub

Private Sub mnulistado_Click()
    Load FrmListadoPromo
    FrmListadoPromo.Show
End Sub

Private Sub Mnumigrar_Click()
    Load FrmMigrar
    FrmMigrar.Show
End Sub

Private Sub mnunotas_Click()
    Load FrmConstancias
    FrmConstancias.Show
End Sub

Private Sub mnusalir_Click()
    Unload Me
End Sub

Private Sub mnuValidar_Click()
    Load FrmValidarFolio
    FrmValidarFolio.Show
End Sub
