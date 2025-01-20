VERSION 5.00
Begin VB.Form FrmActaContinua 
   Caption         =   "Generar Acta Continua"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   14805
   Begin VB.CommandButton CmdGenerar2 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11400
      TabIndex        =   40
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtservidor 
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtpuerto 
      Height          =   375
      Left            =   1920
      TabIndex        =   31
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      Height          =   375
      Left            =   1920
      TabIndex        =   30
      Top             =   8760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtusuario 
      Height          =   375
      Left            =   1920
      TabIndex        =   29
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtde 
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdenviar 
      Caption         =   "enviar"
      Height          =   615
      Left            =   8280
      TabIndex        =   22
      Top             =   9480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdadjunto 
      Caption         =   "adjunto"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtasunto 
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtadjunto 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   8760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtmensaje 
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtpara 
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   39
      Top             =   6960
      Width           =   1200
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13455
      Begin VB.ComboBox CboPensum 
         Height          =   315
         Left            =   8400
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         Height          =   4935
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   12975
         Begin VB.CommandButton cmdrettodos 
            Caption         =   "<<||"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6120
            TabIndex        =   15
            Top             =   3480
            Width           =   495
         End
         Begin VB.CommandButton cmdret 
            Caption         =   "<<"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6120
            TabIndex        =   14
            Top             =   2880
            Width           =   495
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   ">>"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6120
            TabIndex        =   13
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton cmdaddtodos 
            Caption         =   "||>>"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6120
            TabIndex        =   12
            Top             =   1320
            Width           =   495
         End
         Begin VB.ListBox LstSeleccionadas 
            Height          =   4155
            Left            =   6840
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   600
            Width           =   5895
         End
         Begin VB.ListBox LstExistentes 
            Height          =   4155
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   5775
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Seleccionadas"
            Height          =   195
            Left            =   9360
            TabIndex        =   11
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Existentes"
            Height          =   195
            Left            =   2520
            TabIndex        =   10
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.ComboBox CboPeriodo 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox CboTrayecto 
         Height          =   315
         ItemData        =   "FrmActaContinua.frx":0000
         Left            =   4920
         List            =   "FrmActaContinua.frx":0019
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CboPnf 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Pensum:"
         Height          =   195
         Left            =   6840
         TabIndex        =   38
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Academico:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trayecto:"
         Height          =   195
         Left            =   3840
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "P.N.F:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label14 
      Caption         =   "servidor"
      Height          =   255
      Left            =   960
      TabIndex        =   36
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "password"
      Height          =   255
      Left            =   960
      TabIndex        =   34
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "usuario"
      Height          =   255
      Left            =   960
      TabIndex        =   33
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "puerto"
      Height          =   255
      Left            =   960
      TabIndex        =   32
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "De"
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "mensaje"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   9480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Adjunto"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Asunto"
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Para"
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   7920
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmActaContinua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Limpiar_controles(opc As Integer)
    Select Case opc
        Case 1: CboPnf.Text = CboPnf.List(0)
                CboPensum.Clear
                CboPeriodo.Clear
                CboTrayecto.Clear
                LstExistentes.Clear
                LstSeleccionadas.Clear
        Case 2: CboPensum.Text = CboPensum.List(0)
                CboPeriodo.Clear
                CboTrayecto.Clear
                LstExistentes.Clear
                LstSeleccionadas.Clear
        Case 3: CboPeriodo.Text = CboPeriodo.List(0)
                CboTrayecto.Clear
                LstExistentes.Clear
                LstSeleccionadas.Clear
        Case 4: CboTrayecto.Text = CboTrayecto.List(0)
                LstExistentes.Clear
                LstSeleccionadas.Clear
        Case 5: LstExistentes.Clear
                LstSeleccionadas.Clear
    End Select
End Sub

Sub LlenarPNF()
    Dim rs As ADODB.Recordset
    
    funciones.Conectar
    Set rs = cn.Execute("Select despnf from pnf order by despnf asc")
    If Not rs.BOF Then
       CboPnf.Clear
       CboPnf.AddItem ("Seleccione")
       Do While Not rs.EOF
          CboPnf.AddItem rs!despnf
          rs.MoveNext
       Loop
       CboPnf.Text = CboPnf.List(0)
    Else
       MsgBox "No se encontraron pnf"
    End If
    cn.Close
End Sub

Sub LlenarPensum()
    Dim rs As ADODB.Recordset
    
    funciones.Conectar
    Set rs = cn.Execute("select despensum from pensum, pnf where pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & CboPnf.Text & "'")
    If Not rs.BOF Then
       CboPensum.Clear
       CboPensum.AddItem ("Seleccione")
       Do While Not rs.EOF
          CboPensum.AddItem rs!despensum
          rs.MoveNext
       Loop
       CboPensum.Text = CboPensum.List(0)
    Else
       MsgBox "No se encontraron pensum"
    End If
    cn.Close
End Sub

Sub LlenarTrayecto()
    Dim rs As ADODB.Recordset
    Dim maxtrayecto As Integer
    Dim i As Integer
    
    
    funciones.Conectar
    Set rs = cn.Execute("select distinct trayecto from carga_docente where pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "' and periodo='" & CboPeriodo.Text & "' order by trayecto asc")
    If Not rs.BOF Then
       CboTrayecto.Clear
       CboTrayecto.AddItem ("Seleccione")
       Do While Not rs.EOF
        Select Case rs!trayecto
                 Case "01"
                     CboTrayecto.AddItem "Inicial"
                 Case "03", "04"
                     CboTrayecto.AddItem "Transicion"
                     
                 Case Else
                     CboTrayecto.AddItem rs!trayecto
        End Select
        rs.MoveNext
       Loop
       CboTrayecto.Text = CboPnf.List(0)
    Else
       MsgBox "No se encontraron trayectos"
    End If
    cn.Close

End Sub

Sub LlenarPeriodos()
    Dim rs As ADODB.Recordset
    
    funciones.Conectar
    'Set rs = cn.Execute("Select distinct periodo from carga_docente where pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "'")
    Set rs = cn.Execute("select distinct periodo, cast(substring(periodo  from 3 for 4) as int) as orden1, cast(substring(periodo  from 1 for 1) as int) as orden2" & _
          " FROM carga_docente where pnf='" & CboPnf.Text & "' and periodo<>'ACTUAL' order by orden1 desc, orden2 desc")

    If Not rs.BOF Then
       CboPeriodo.Clear
       CboPeriodo.AddItem ("Seleccione")
       Do While Not rs.EOF
          CboPeriodo.AddItem rs!periodo
          rs.MoveNext
       Loop
       CboPeriodo.Text = CboPeriodo.List(0)
    Else
       MsgBox "No se encontraron periodos"
    End If
    cn.Close
End Sub


Private Sub CboPensum_Click()
    If CboPensum.Text <> "" Then
        If CboPensum.Text <> "Seleccione" Then
            Limpiar_controles (3)
            LlenarPeriodos
        End If
    End If
End Sub

Private Sub CboPeriodo_Click()
    If CboPeriodo.Text <> "" Then
        If CboPeriodo.Text <> "Seleccione" Then
            Limpiar_controles (5)
            
            funciones.Conectar
            If FrmActaContinua.Caption = "Generar Acta Continua" Then
                Set rs = cn.Execute("select distinct facilitador from carga_docente where  facilitador <>'' and pnf='" & CboPnf.Text & "' and periodo='" & CboPeriodo.Text & "' order by facilitador asc")
            Else
                'Set rs = cn.Execute("select codigo,unidad,seccion from carga_docente_per where  pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "' and periodo='" & CboPeriodo.Text & "' and trayecto='" & trayecto & "' order by unidad asc")
            End If
            If Not rs.BOF Then
               LstExistentes.Clear
               
               Do While Not rs.EOF
                  LstExistentes.AddItem rs!facilitador
                  rs.MoveNext
               Loop
               cmdaddtodos.Enabled = True
            Else
               MsgBox "No existen secciones para el P.N.F, en este periodo"
            End If
            cn.Close
        End If
    End If
End Sub

Private Sub CboPnf_Click()
    If CboPnf.Text <> "" Then
        If CboPnf.Text <> "Seleccione" Then
            Limpiar_controles (3)
            LlenarPeriodos
        End If
    End If
End Sub

Private Sub CboTrayecto_Click()
Dim rs As ADODB.Recordset
Dim trayecto As String

    If CboTrayecto.Text <> "" Then
        If CboTrayecto.Text <> "Seleccione" Then
            Limpiar_controles (5)
            Select Case CboTrayecto.Text
                Case "Inicial"
                    trayecto = "01"
                Case "Transicion"
                    funciones.Conectar
                    
                    Set rs = cn.Execute("select max(pensumnivel.trapensumnivel) as transicion from pensumnivel, pensum, pnf where pensumnivel.pensum_codpensum=pensum.codpensum and" & _
                        " pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & CboPnf.Text & "' and pensum.despensum='" & CboPensum.Text & "'")
                    If Not rs.BOF Then
                        trayecto = "0" & rs!transicion
                    End If
                Case Else
                    trayecto = CboTrayecto.Text
            End Select
            funciones.Conectar
            If FrmActaContinua.Caption = "Generar Acta Continua" Then
                Set rs = cn.Execute("select codigo,unidad,seccion from carga_docente where  pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "' and periodo='" & CboPeriodo.Text & "' and trayecto='" & trayecto & "' order by unidad asc")
            Else
                Set rs = cn.Execute("select codigo,unidad,seccion from carga_docente_per where  pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "' and periodo='" & CboPeriodo.Text & "' and trayecto='" & trayecto & "' order by unidad asc")
            End If
            If Not rs.BOF Then
               LstExistentes.Clear
               
               Do While Not rs.EOF
                  LstExistentes.AddItem rs!codigo & "-" & rs!unidad & " [" & rs!seccion & "]"
                  rs.MoveNext
               Loop
               cmdaddtodos.Enabled = True
            Else
               MsgBox "No existen secciones para el trayecto, en este periodo"
            End If
            cn.Close
        End If
    End If
End Sub

Private Sub cmdadd_Click()
    Call LstExistentes_DblClick
End Sub

Private Sub cmdaddtodos_Click()
    For i = 0 To LstExistentes.ListCount - 1
        LstSeleccionadas.AddItem (LstExistentes.List(i))
    Next
    LstExistentes.Clear
    cmdrettodos.Enabled = True
    cmdaddtodos.Enabled = False
    CmdGenerar2.Enabled = True
End Sub

Private Sub cmdadjunto_Click()
    'With CommonDialog1
    '    .ShowOpen
    '    If .FileName = "" Then
    '        Exit Sub
    '    End If
    '    txtadjunto.Text = .FileName
    'End With
    
End Sub


     'Option Explicit
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    ' El ejemplo para poder enviar el mail necesita la referencia a: _
      > Miscrosoft CDO Windows For 2000 Library ( es el archivo dll cdosys.dll )
      
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
      
   
      
Private Sub Command1_Click()

End Sub

Private Sub CmdGenerar2_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim materia As String
    Dim seccion As String
    Dim trayecto As String
    Dim contador As Integer
    Dim sql As String
    
    funciones.Conectar
    
    Dim fs, f, s
    'define la ruta para crear la carpeta en el Escritorio con los parámetros seleccionados
    folderspec = "C:\Users\USUARIO\Desktop\" & CboPnf.Text & " " & CboPeriodo.Text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.folderexists(folderspec) Then
        Set f = fs.CreateFolder(folderspec)
    End If
    
    For i = 0 To LstSeleccionadas.ListCount - 1
        'define la ruta para crear la carpeta en el Escritorio con los parámetros seleccionados
        folderspec = "C:\Users\USUARIO\Desktop\" & CboPnf.Text & " " & CboPeriodo.Text & "\" & LstSeleccionadas.List(i)
                    
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Not fs.folderexists(folderspec) Then
            Set f = fs.CreateFolder(folderspec)
        End If
        
        Set rs = cn.Execute("select cedula,facilitador,codigotrayecto,seccion,codigo,unidad From carga_docente where pnf='" & CboPnf.Text & "' and periodo='" & CboPeriodo.Text & "' and facilitador='" & LstSeleccionadas.List(i) & "'")
        If Not rs.BOF Then
            Do While Not rs.EOF
                 Set AppExcel = CreateObject("Excel.application")
                 If FrmActaContinua.Caption = "Generar Acta Continua" Then
                     AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "Plantillas\plantilla2_con_Asistencia.xlsm"  'para abrir el libro
                     AppExcel.Application.Windows("plantilla2_con_Asistencia.xlsm").Activate
                 Else
                     AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "Plantillas\plantilla_per.xls"  'para abrir el libro
                     AppExcel.Application.Windows("plantilla_per.xls").Activate
                 End If
                 
                 'AppExcel.application.Visible = True
                 AppExcel.Application.Cells(7, 1).Value = CboPnf.Text
                 AppExcel.Application.Cells(9, 1).Value = rs!codigo & "-" & rs!unidad & "(" & rs!seccion & ")"
                 AppExcel.Application.Cells(7, 7).Value = CboPeriodo.Text
                 AppExcel.Application.Cells(11, 1).Value = LstSeleccionadas.List(i)
                 AppExcel.Application.Cells(11, 7).Value = rs!cedula
                 'proporciona el nombre al archivo
                 archivo = rs!codigo & "-" & rs!unidad & "-" & rs!seccion
                 
                 sql = "select minaprobatoria as resultado from excepcioncalificacion where escalaevaluacion_codescalaevaluacion=2 and unidadcurricular_codunidadcurricular='" & materia & "'"
                 minaprob = funciones.CampoEntero(sql, cn)
                 If minaprob > 0 Then
                     AppExcel.Application.Cells(12, 2).Value = minaprob
                 Else
                     AppExcel.Application.Cells(12, 2).Value = 12
                 End If
                
                 If FrmActaContinua.Caption = "Generar Acta Continua" Then
                     Set rs2 = cn.Execute("select cedulapar as cedula,nombrepar as nombres,apellidopar as apellidos from actanotas_continua where cedulafac='" & rs!cedula & "' and periodo='" & CboPeriodo.Text & "' and pnf='" & CboPnf.Text & "' and codigou='" & rs!codigo & "' and seccion='" & rs!seccion & "' order by apellidopar, nombrepar")
                 Else
                     ' Set rs = cn.Execute("select distinct usuario.cedusuario as cedula, usuario.apeusuario as apellidos, usuario.nomusuario as nombres from secciones, unidadcurricular, inscripcionucurricular, participante, usuario, periodosacademicos, pnf, pensumunidadcurricular, pensum, inscripcion_per" & _
                     '" where unidadcurricular.codunidadcurricular=secciones.unidadcurricular_codunidadcurricular and inscripcionucurricular.secciones_codsecciones=secciones.codsecciones and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
                     '" pensumunidadcurricular.pensum_codpensum=pensum.codpensum and pnf.codpnf=unidadcurricular.pnf_codpnf and inscripcionucurricular.periodosacademicos_codperiodosacademicos=periodosacademicos.codperiodosacademicos and inscripcionucurricular.participantes_codparticipantes=participante.codparticipantes and" & _
                     '" inscripcionucurricular.codinscripcionucurricular=inscripcion_per.inscripcionucurricular_codinscripcionucurricular and" & _
                     '" participante.usuario_cedusuario=usuario.cedusuario and unidadcurricular.codunidadcurricular='" & materia & "' and secciones.nomsecciones='" & seccion & "' and periodosacademicos.desperiodosacademicos='" & CboPeriodo.Text & "' and pnf.despnf='" & _
                     'CboPnf.Text & "' and unidadcurricular.trayecto_codtrayecto =" & rs!codigotrayecto & " and pensum.despensum='" & CboPensum.Text & "' order by apellidos, nombres asc")
                 End If
                 If Not rs2.BOF Then
                     fila = 15
                     contador = 1
                     Do While Not rs2.EOF
                         AppExcel.Application.Cells(fila, 1).Value = contador
                         AppExcel.Application.Cells(fila, 2).Value = rs2!cedula
                         AppExcel.Application.Cells(fila, 3).Value = rs2!apellidos
                         AppExcel.Application.Cells(fila, 4).Value = rs2!nombres
                         rs2.MoveNext
                         contador = contador + 1
                         fila = fila + 1
                     Loop
                     fila = fila + 2
                      AppExcel.Application.Cells(fila, 2).Value = "Vocero:"
                      AppExcel.Application.Cells(fila + 1, 3).Value = "Firma"
                      AppExcel.Application.Cells(fila + 1, 4).Value = "_________________"
                      AppExcel.Application.Cells(fila + 2, 3).Value = "Apellidos y Nombres"
                      AppExcel.Application.Cells(fila + 2, 4).Value = "_________________"
                      AppExcel.Application.Cells(fila + 3, 3).Value = "Cedula"
                      AppExcel.Application.Cells(fila + 3, 4).Value = "_________________"
                      AppExcel.Application.Cells.Columns("B:D").EntireColumn.AutoFit
                      'archivo = "AAA"
                      AppExcel.Application.ActiveWorkbook.SaveAs FileName:=folderspec & "\" & archivo & ".xlsm"
                      AppExcel.Application.Windows(archivo & ".xlsm").Close
                 End If
                 rs.MoveNext
            Loop
            'CommonDialog1.ShowOpen
            'If CommonDialog1.FileName <> "" Then
            '    ruta = CommonDialog1.FileName
            'Else
                'ruta = "C:\Documents and Settings\Administrador\Escritorio\a\"
                
                'ruta = "C:\Documents and Settings\usuario\Escritorio\a\"
            'End If
            'AppExcel.application.Visible = True
            'archivo = Replace(LstSeleccionadas.List(i), "[", "(")
            'archivo = Replace(archivo, "]", ")")
            'AppExcel.Application.ActiveWorkbook.SaveAs FileName:=folderspec & "\" & archivo & ".xlsm"
            'AppExcel.application.Windows(LstSeleccionadas.List(i) & ".xls").Close
            
            
            'AppExcel.Application.Saved = True
            AppExcel.Application.Quit
            Set AppExcel = Nothing
        End If
        
    Next
    cn.Close
    MsgBox "Archivos generados en: " & folderspec & Chr(13) & "Proceso Concluido"

End Sub

    Private Sub Form_Load()
        LlenarPNF
        'Me.Caption = " Ejemplo para enviar correo usando la libreria Microsoft CDO "
        cmdenviar.Caption = " Enviar mail "
          
        txtservidor.Text = "smtp.gmail.com"
        txtpara = "rinconjf@hotmail.com.com"
        txtde = "rinconjf@gmail.com"
        txtasunto = "Prueba"
        txtmensaje = " ... Cuerpo del mensaje "
        txtadjunto = vbNullString
        txtpuerto.Text = 465
        txtpassword = "sebastian"
        txtusuario = "rinconjf@gmail.com"
    End Sub
   

Private Sub cmdgenerar_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim materia As String
    Dim seccion As String
    Dim trayecto As String
    Dim contador As Integer
    Dim celda As String
    
    Dim fs, f, s
    'define la ruta para crear la carpeta en el Escritorio con los parámetros seleccionados
    folderspec = "C:\Users\USUARIO\Desktop\" & CboPnf.Text & " trayecto " & CboTrayecto.Text & " " & CboPeriodo.Text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.folderexists(folderspec) Then
        Set f = fs.CreateFolder(folderspec)
    End If
            
    funciones.Conectar
    For i = 0 To LstSeleccionadas.ListCount - 1
        materia = Mid(LstSeleccionadas.List(i), 1, InStr(1, LstSeleccionadas.List(i), "-") - 1)
        seccion = Mid(LstSeleccionadas.List(i), InStr(1, LstSeleccionadas.List(i), " (") + 2, (InStr(1, LstSeleccionadas.List(i), ")") - 2) - (InStr(1, LstSeleccionadas.List(i), " (")))
        'Set rs = cn.Execute("select distinct usuario.cedusuario as cedula, usuario.nomapeusuario as nomcompleto from pnf,secciones, unidadcurricular, cargaacademica, cargahoraria, facilitador, usuario, periodosacademicos" & _
        '    " where unidadcurricular.codunidadcurricular=secciones.unidadcurricular_codunidadcurricular and cargaacademica.secciones_codsecciones=secciones.codsecciones and unidadcurricular.pnf_codpnf=pnf.codpnf and" & _
        '    " cargaacademica.cargahoraria_codcargahoraria=cargahoraria.codcargahoraria and cargahoraria.facilitador_codfacilitador=facilitador.codfacilitador and facilitador.usuario_cedusuario=usuario.cedusuario and" & _
        '    " secciones.periodosacademicos_codperiodoacademico=periodosacademicos.codperiodosacademicos and unidadcurricular.codunidadcurricular='" & materia & "' and secciones.nomsecciones='" & seccion & "' and" & _
        '    " pnf.despnf='" & CboPnf.Text & "' and periodosacademicos.desperiodosacademicos='" & CboPeriodo.Text & "'")
        Set rs = cn.Execute("select cedula,facilitador,trayecto from carga_docente where pnf='" & CboPnf.Text & "' and pensum='" & CboPensum.Text & "' and codigo='" & materia & "' and seccion='" & seccion & "' and periodo='" & CboPeriodo.Text & "'")
        If Not rs.BOF Then
            Set AppExcel = CreateObject("Excel.application")
            AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "Plantillas\plantilla2.xlsx"  'para abrir el libro
            AppExcel.Application.Windows("plantilla2.xlsX").Activate
            'AppExcel.application.Visible = True
            AppExcel.Application.Cells(7, 1).Value = CboPnf.Text
            AppExcel.Application.Cells(9, 1).Value = LstSeleccionadas.List(i)
            AppExcel.Application.Cells(8, 7).Value = CboPeriodo.Text
            AppExcel.Application.Cells(11, 1).Value = rs!facilitador
            AppExcel.Application.Cells(11, 7).Value = rs!cedula
            trayecto = rs!trayecto
           'Select Case rs!trayecto
           '     Case "Inicial"
           '         trayecto = "01"
           '     Case "Transicion"
           '         funciones.Conectar
           '
           '         Set rs = cn.Execute("select max(pensumnivel.trapensumnivel) as transicion from pensumnivel, pensum, pnf where pensumnivel.pensum_codpensum=pensum.codpensum and" & _
           '             " pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & CboPnf.Text & "' and pensum.despensum='" & CboPensum.Text & "'")
           '         If Not rs.BOF Then
           '             trayecto = "0" & rs!transicion
           '         End If
           '     Case Else
           '         trayecto = CboTrayecto.Text
           ' End Select
            Set rs = cn.Execute("select distinct usuario.cedusuario as cedula, usuario.apeusuario as apellidos, usuario.nomusuario as nombres from secciones, unidadcurricular, inscripcionucurricular, participante, usuario, periodosacademicos, pnf, pensumunidadcurricular, pensum" & _
                " where unidadcurricular.codunidadcurricular=secciones.unidadcurricular_codunidadcurricular and inscripcionucurricular.secciones_codsecciones=secciones.codsecciones and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
                " pensumunidadcurricular.pensum_codpensum=pensum.codpensum and pnf.codpnf=unidadcurricular.pnf_codpnf and inscripcionucurricular.periodosacademicos_codperiodosacademicos=periodosacademicos.codperiodosacademicos and inscripcionucurricular.participantes_codparticipantes=participante.codparticipantes and" & _
                " participante.usuario_cedusuario=usuario.cedusuario and unidadcurricular.codunidadcurricular='" & materia & "' and secciones.nomsecciones='" & seccion & "' and periodosacademicos.desperiodosacademicos='" & CboPeriodo.Text & "' and pnf.despnf='" & _
                CboPnf.Text & "' and unidadcurricular.traunidadcurricular ='" & rs!trayecto & "' and pensum.despensum='" & CboPensum.Text & "' order by apellidos, nombres asc")
            If Not rs.BOF Then
                fila = 14
                contador = 1
                Do While Not rs.EOF
                    AppExcel.Application.Cells(fila, 1).Value = contador
                    AppExcel.Application.Cells(fila, 2).Value = rs!cedula
                    AppExcel.Application.Cells(fila, 3).Value = rs!apellidos
                    AppExcel.Application.Cells(fila, 4).Value = rs!nombres
                    celda = "E" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    AppExcel.Application.Selection.FormulaHidden = False
                    celda = "F" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-1]*0.2"
                    celda = "G" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    AppExcel.Application.Selection.FormulaHidden = False
                    celda = "H" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-1]*0.2"
                    celda = "I" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    AppExcel.Application.Selection.FormulaHidden = False
                    celda = "J" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-1]*0.2"
                    celda = "K" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    AppExcel.Application.Selection.FormulaHidden = False
                    celda = "L" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-1]*0.2"
                    celda = "M" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    AppExcel.Application.Selection.FormulaHidden = False
                    celda = "N" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-1]*0.2"
                    celda = "O" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=RC[-9]+RC[-7]+RC[-5]+RC[-3]+RC[-1]"
                    celda = "P" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.activecell.FormulaR1C1 = "=IF((RC[-1]-TRUNC(RC[-1]))<0.44999999999,TRUNC(RC[-1]),TRUNC(RC[-1]) +1)"
                    celda = "Q" & fila
                    AppExcel.Application.Range(celda).Select
                    AppExcel.Application.Selection.Locked = False
                    rs.MoveNext
                    contador = contador + 1
                    fila = fila + 1
                Loop
                fila = fila + 1
                 AppExcel.Application.Cells(fila, 2).Value = "Vocero:"
                 AppExcel.Application.Cells(fila + 1, 3).Value = "Firma"
                 AppExcel.Application.Cells(fila + 1, 4).Value = "_________________"
                 AppExcel.Application.Cells(fila + 2, 3).Value = "Apellidos y Nombres"
                 AppExcel.Application.Cells(fila + 2, 4).Value = "_________________"
                 AppExcel.Application.Cells(fila + 3, 3).Value = "Cedula"
                 AppExcel.Application.Cells(fila + 3, 4).Value = "_________________"
                 AppExcel.Application.Sheets("Evaluacion Continua").Select
                 'AppExcel.Application.Sheets("Evaluacion Continua").Protect Password:="@dm1nC333"
                 AppExcel.Application.ActiveSheet.Protect Password:="@dm1nC333", DrawingObjects:=True, Contents:=True, Scenarios:=True
            End If
            
            
            
            
            
            

            'CommonDialog1.ShowOpen
            'If CommonDialog1.FileName <> "" Then
            '    ruta = CommonDialog1.FileName
            'Else
                'ruta = "C:\Documents and Settings\Administrador\Escritorio\a\"
                'ruta = "C:\Users\USUARIO\Desktop\a\"
                'ruta = "C:\Documents and Settings\usuario\Escritorio\a\"
            'End If
            'AppExcel.application.Visible = True
            AppExcel.Application.ActiveWorkbook.SaveAs FileName:=folderspec & "\" & LstSeleccionadas.List(i) & ".xlsx"
            'AppExcel.application.Windows(LstSeleccionadas.List(i) & ".xls").Close
            AppExcel.Application.Quit
            Set AppExcel = Nothing
        End If
        
    Next
    cn.Close
    MsgBox "Archivos generados en: " & folderspec & Chr(13) & "Proceso Concluido"
End Sub

Private Sub cmdret_Click()
    Call LstSeleccionadas_DblClick
End Sub

Private Sub cmdrettodos_Click()
    For i = 0 To LstSeleccionadas.ListCount - 1
        LstExistentes.AddItem (LstSeleccionadas.List(i))
    Next
    LstSeleccionadas.Clear
    cmdrettodos.Enabled = False
    cmdgenerar.Enabled = False
End Sub

Private Sub LstExistentes_Click()
    'If LstExistentes.Selected Then
        cmdadd.Enabled = True
        cmdrettodos.Enabled = True
    'End If
End Sub

Private Sub LstExistentes_DblClick()
    LstSeleccionadas.AddItem LstExistentes.Text
    LstExistentes.RemoveItem (LstExistentes.ListIndex)
    cmdadd.Enabled = False
    If LstSeleccionadas.ListCount > 0 Then
        CmdGenerar2.Enabled = True
    Else
        CmdGenerar2.Enabled = False
    End If
End Sub

Private Sub LstSeleccionadas_Click()
    'If LstSeleccionadas.Selected Then
        cmdret.Enabled = True
    'End If
End Sub

Private Sub LstSeleccionadas_DblClick()
    LstExistentes.AddItem LstSeleccionadas.Text
    LstSeleccionadas.RemoveItem (LstSeleccionadas.ListIndex)
    cmdret.Enabled = False
    cmdrettodos.Enabled = False
    If LstSeleccionadas.ListCount > 0 Then
        CmdGenerar2.Enabled = True
    Else
        CmdGenerar2.Enabled = False
    End If
End Sub
