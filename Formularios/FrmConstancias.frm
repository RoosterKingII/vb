VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmConstancias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Notas Certificadas de Graduado"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12855
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   12855
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   615
      Left            =   5640
      TabIndex        =   30
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3360
      TabIndex        =   27
      Top             =   9600
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Notas"
      Height          =   5055
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   12615
      Begin MSFlexGridLib.MSFlexGrid FlexNotas 
         Height          =   4695
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   7
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Materias"
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   12615
      Begin VB.TextBox TxtSeccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10320
         MaxLength       =   4
         TabIndex        =   28
         Top             =   1080
         Width           =   600
      End
      Begin VB.ComboBox CboMateria 
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   1080
         Width           =   7215
      End
      Begin VB.ComboBox CboPeriodo 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1470
      End
      Begin VB.TextBox TxtNota 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11760
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1080
         Width           =   600
      End
      Begin VB.ComboBox CboSemestre 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sección:"
         Height          =   195
         Left            =   9600
         TabIndex        =   29
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Semestre:"
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Materia:"
         Height          =   195
         Left            =   4920
         TabIndex        =   23
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         Height          =   195
         Left            =   11280
         TabIndex        =   21
         Top             =   1200
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Participante"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12615
      Begin VB.TextBox txtcedula 
         Height          =   375
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtapellidos 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtnombres 
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox cbopnf 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtcohorte 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cbopensum 
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtespecialidad 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtpensum 
         Height          =   315
         Left            =   6120
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cédula:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
         Height          =   195
         Left            =   5160
         TabIndex        =   13
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "P.N.F:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Pensum:"
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cohorte:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdnotascertificadas 
      Caption         =   "Imprimir Notas Certificadas"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7800
      TabIndex        =   1
      Top             =   9600
      Width           =   2295
   End
End
Attribute VB_Name = "FrmConstancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboElectiva_Click()
    TxtNotaElectiva.SetFocus
End Sub

Private Sub CboMateria_Click()
    TxtSeccion.SetFocus
End Sub

Private Sub CboPensum_Click()
    If CboPensum.Text <> "Seleccione" Then
        If Len(CboPensum.Text) > 0 Then
            txtpensum.Text = CboPensum.Text
            txtpensum.SetFocus
        End If
    End If
End Sub

Private Sub CboPeriodo_Click()
    Dim sql As String
    If Len(CboPeriodo.Text) > 0 Then
        If CboPeriodo.Text <> "Seleccione" Then
            funciones.Conectar
            sql = "select desunidadcurricular as resultado from unidadcurricular, pensumunidadcurricular,pensum,pnf where unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and pensumunidadcurricular.pensum_codpensum=pensum.codpensum and" & _
                  " pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and unidadcurricular.traunidadcurricular='" & CboSemestre.Text & "' and pensum.despensum='" & txtpensum.Text & "' order by resultado asc"
            funciones.llenarcombobox CboMateria, sql, cn, True
            cn.Close
        End If
    End If
End Sub

Private Sub CboPnf_Click()
    If CboPnf.Text <> "Seleccione" Then
        If Len(CboPnf.Text) > 0 Then
            txtespecialidad.Text = CboPnf.Text
            txtespecialidad.SetFocus
        End If
    End If
End Sub

Private Sub CboSemestre_Click()
    Dim sql As String
    
    sql = "select desperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos<>'ACTUAL' order by cast(substring(desperiodosacademicos from 3 for 4) as  int), cast(substring(desperiodosacademicos from 1 for 1) as int) desc"
    funciones.Conectar
    funciones.llenarcombobox CboPeriodo, sql, cn, True
    sql = "select desunidadcurricular as resultado from unidadcurricular, pensumunidadcurricular,pensum,pnf where unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and pensumunidadcurricular.pensum_codpensum=pensum.codpensum and" & _
          " pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and unidadcurricular.traunidadcurricular='" & CboSemestre.Text & "' and pensum.despensum='" & txtpensum.Text & "' order by resultado asc"
    funciones.llenarcombobox CboMateria, sql, cn, True
    cn.Close
End Sub

Private Sub CmdGuardar_Click()
    Dim codperiodo As Integer
    Dim codseccion As Integer
    Dim codpnf As Integer
    Dim codpensum As Integer
    Dim codunidadcurricular As String
    Dim trayecto As String
    Dim turno As Integer
    Dim codparticipante As Long
    Dim sql As String
    Dim anoper As Integer
    Dim numper As Integer
    Dim codescala As Integer
    Dim condicion As String
    
    funciones.Conectar
    For i = 1 To FlexNotas.Rows - 1
        sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & txtcedula.Text & "'"
        codparticipante = funciones.CampoEnteroLargo(sql, cn)
        sql = "select codperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos='" & FlexNotas.TextMatrix(i, 5) & "'"
        codperiodo = funciones.CampoEntero(sql, cn)
        sql = "select codpnf as resultado from pnf where despnf='" & txtespecialidad.Text & "'"
        codpnf = funciones.CampoEntero(sql, cn)
        sql = "select pensum.codpensum as resultado from pensum, cohorte where pensum.codpensum=cohorte.pensum_codpensum and  cohorte.pnf_codpnf=" & codpnf & " and pensum.despensum='" & txtpensum.Text & "' and cohorte.descohorte='" & txtcohorte.Text & "'"
        codpensum = funciones.CampoEntero(sql, cn)
        trayecto = numerotrayecto(FlexNotas.TextMatrix(i, 6), codpensum, codpnf, txtcohorte.Text)
        sql = "select codunidadcurricular as resultado from unidadcurricular where desunidadcurricular='" & FlexNotas.TextMatrix(i, 1) & "' and traunidadcurricular='" & trayecto & "' and pnf_codpnf=" & codpnf
        codunidadcurricular = funciones.CampoString(sql, cn)
        sql = "select codsecciones as resultado from secciones where periodosacademicos_codperiodoacademico=" & codperiodo & " and unidadcurricular_codunidadcurricular='" & codunidadcurricular & "' and nomsecciones='" & FlexNotas.TextMatrix(i, 3) & "'"
        codseccion = funciones.CampoEntero(sql, cn)
        'si no existe la seccion
        If codseccion = 0 Then
            'se crea la seccion
            sql = "select  max (codsecciones) as resultado from secciones"
            codseccion = funciones.proximocodigoregistro(sql, cn)
            'se determina el turno
            Select Case Mid(FlexNotas.TextMatrix(i, 3), 1, 2)
                Case "01": turno = 1
                Case "02": turno = 2
                Case "03": turno = 3
            Case Else
                turno = 1
            End Select
            sql = "insert into secciones values(" & codseccion & "," & turno & ",'" & FlexNotas.TextMatrix(i, 3) & "',40," & codperiodo & ",'" & codunidadcurricular & "','false')"
            cn.Execute (sql)
            'se crea la inscripcion_seccion
            sql = "insert into inscripcion_seccion values(" & codseccion & ",0,0)"
            cn.Execute (sql)
        End If
        sql = "select codinscripcionucurricular as resultado from inscripcionucurricular,participante where participante.codparticipantes=inscripcionucurricular.participantes_codparticipantes and " & _
              " inscripcionucurricular.unidadcurricular_codunidadcurricular='" & codunidadcurricular & "' and secciones_codsecciones=" & codseccion & " and periodosacademicos_codperiodosacademicos=" & codperiodo & " and participante.usuario_cedusuario='" & txtcedula.Text & "'"
        'si no existe inscripcion para el periodoacademico, unidadcurricular y seccion
        If Not funciones.ExisteRegistro(sql, cn) Then
            sql = "select max(codinscripcionucurricular) as resultado from inscripcionucurricular"
            codinscripcion = funciones.CampoEnteroLargo(sql, cn) + 1
            numper = CInt(Mid(FlexNotas.TextMatrix(i, 5), 1, 1))
            anoper = CInt(Mid(FlexNotas.TextMatrix(i, 5), 3, 4))
            If anoper <= 2012 And numper < 6 Then
                codescala = 1
            Else
                codescala = 2
            End If
            'se determina la condicion
            If codescala = 1 Then
                If CInt(FlexNotas.TextMatrix(i, 2)) >= 10 Then
                    condicion = "Aprobado"
                Else
                    condicion = "Reprobado"
                End If
            Else
                'se verifica si la unidad curricular es una excepcion
                sql = "select minaprobatoria as resultado from excepcioncalificacion where unidadcurricular_codunidadcurricular='" & codunidadcurricular & "' and escalaevaluacion_codescalaevaluacion=2"
                If funciones.ExisteRegistro(sql, cn) Then
                    If CInt(FlexNotas.TextMatrix(i, 2)) >= funciones.CampoEntero(sql, cn) Then
                        condicion = "Aprobado"
                    Else
                        condicion = "Reprobado"
                    End If
                Else    'en caso de no serlo
                    If CInt(FlexNotas.TextMatrix(i, 2)) >= 12 Then
                        condicion = "Aprobado"
                    Else
                        condicion = "Reprobado"
                    End If
                End If
            End If
            'procedemos a inscribirlo
            sql = "insert into inscripcionucurricular values (" & codperiodo & "," & codinscripcion & ",'" & codunidadcurricular & "'," & codparticipante & "," & _
                  codescala & "," & codseccion & ",'" & FlexNotas.TextMatrix(i, 2) & "','" & condicion & "')"
            cn.Execute (sql)
        End If
    Next i
    
End Sub

Private Sub CmdLimpiar_Click()
    Dim i As Integer
    
    txtcedula.Text = ""
    txtapellidos.Text = ""
    txtnombres.Text = ""
    CboPnf.Clear
    txtespecialidad.Text = ""
    CboPensum.Clear
    txtpensum.Text = ""
    txtcohorte.Text = ""
    CboSemestre.Clear
    CboPeriodo.Clear
    CboMateria.Clear
    TxtSeccion.Text = ""
    TxtNota.Text = ""
    For i = 0 To FlexNotas.Cols - 1
        FlexNotas.TextMatrix(1, i) = ""
    Next
    FlexNotas.Rows = 2
    txtcedula.SetFocus
End Sub

Private Sub cmdnotascertificadas_Click()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim fila As Integer
    Dim pagina As Integer
    Dim control As String
    Dim codpnf As Integer
    Dim codpromocion As Integer
    Dim condicion As String
    Dim pag As Integer
    'inicializa el total de creditos cursados
    totaluccur = 0
    totalnotauc = 0
    uc = 0
    ucxnota = 0
    pagina = 1
    pag = 1
    res = 1
    fila = 7
    FrmConstancias.MousePointer = vbHourglass
    funciones.Conectar
    
    sql = "select codpensum as resultado from pensum, pnf where pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and despensum='" & txtpensum.Text & "'"
    codpensum = funciones.CampoEntero(sql, cn)
    sql = "select codpnf as resultado from pnf where despnf='" & txtespecialidad.Text & "'"
    codpnf = funciones.CampoEntero(sql, cn)
    sql = "select id as resultado from solicitudes where cedusuario='" & txtcedula.Text & "' and tiposolicitud_codtiposolicitud=7 and estatussolicitud_codestatussolicitud<=2"
    If Not funciones.ExisteRegistro(sql, cn) Then   'inserta la solicitud y calcula el numero de control
        sql = "select count (id) as resultado from solicitudes where tiposolicitud_codtiposolicitud=7 and (fechasolicitud >='" & Year(Date) & "-" & Month(Date) & "-01' and fechasolicitud<='" & funciones.fin_del_Mes(Date) & "')"
        control = formatocorrelativo(funciones.CampoEntero(sql, cn) + 1)
        control = Year(Date) & "-" & Month(Date) & "-" & control
        sql = "select max(id) as resultado from solicitudes"
        codsolicitud = funciones.proximocodigoregistro(sql, cn)
        sql = "insert into solicitudes values(" & codsolicitud & ",'" & txtcedula.Text & "','',7,2,'" & Format(Date, "yyyy-mm-dd") & "',1,'" & control & "','')"
        cn.Execute (sql)
    Else    'actualiza el estatus de la solicitud
        sql = "select numerosolicitud as resultado from solicitudes where cedusuario='" & txtcedula.Text & "' and tiposolicitud_codtiposolicitud=7 and estatussolicitud_codestatussolicitud<=2"
        control = funciones.CampoString(sql, cn)
        If control = "000" Then
            sql = "select count (id) as resultado from solicitudes where tiposolicitud_codtiposolicitud=7 and (fechasolicitud >='" & Year(Date) & "-" & Month(Date) & "-01' and fechasolicitud<='" & funciones.fin_del_Mes(Date) & "')"
            control = formatocorrelativo(funciones.CampoEntero(sql, cn) + 1)
            control = Year(Date) & "-" & Month(Date) & "-" & control
            sql = "select max(id) as resultado from solicitudes"
            codsolicitud = funciones.proximocodigoregistro(sql, cn)
            sql = "insert into solicitudes values(" & codsolicitud & ",'" & txtcedula.Text & "','',7,2,'" & Format(Date, "yyyy-mm-dd") & "',1,'" & control & "','')"
            cn.Execute (sql)
        End If
    End If
    'crea el objeto de excel y abre la aplicacion
    Set AppExcel = CreateObject("Excel.application")
    AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "PLANTILLAS\Certificadas.xls"  'para abrir el libro
    AppExcel.Application.Windows("Certificadas.XLS").Activate
    'AppExcel.Application.Visible = True
    sql = "select codpromocion as resultado from autenticidad_titulo where cedula='" & txtcedula.Text & "' and pnf='" & txtespecialidad.Text & "'"
    If funciones.ExisteRegistro(sql, cn) Then   'si esta registrado en una promocion
        sql = "select trayectoini as trapensumnivel, trayectofin as finpensumnivel,coorte as descohorte,codpromocion,codparticipantes,nivel as codnivel from autenticidad_titulo where cedula='" & txtcedula.Text & "' and codpnf=" & codpnf & " order by codnivel asc"
        Set rs = cn.Execute(sql)
        If Not rs.BOF Then
            
            Do While Not rs.EOF
                If rs!codnivel = 1 Then
                    trainicio = rs!trapensumnivel
                Else
                    sql = "select pensumnivel.trapensumnivel as resultado from participantecohorte, pensumnivel where participantecohorte.nivel_codnivel=pensumnivel.nivel_codnivel and pensumnivel.pensum_codpensum=" & codpensum & " and pensumnivel.nivel_codnivel=3"
                    trainicio = funciones.CampoEntero(sql, cn)
                    
                End If
                transicion = rs!finpensumnivel
                codparticipante = rs!codparticipantes
                codpromocion = rs!codpromocion
                condicion = "' and cast(traunidadcurricular as int)>=" & trainicio & " and cast(traunidadcurricular as int)<=" & transicion
                If (Mid(rs!descohorte, 1, 2) = "IN") And trainicio > 1 Then
                    condicion = condicion & " and traunidadcurricular<>'0" & trainicio & "'"
                Else
                    'aqui se consulta si es egresado de pnf para omitir el trayecto de transicion
                    If cn.State = 0 Then
                        funciones.Conectar
                        
                    End If
                    
                    sql = "select codparticipantecohorte as resultado from participantecohorte, cohorte where participantecohorte.participante_codparticipante=" & codparticipante & " and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and substring(cohorte.descohorte from 1 for 2)='IN'"
                    If (Mid(rs!descohorte, 1, 2) = "PS" And funciones.ExisteRegistro(sql, cn)) Then
                        condicion = condicion & " and traunidadcurricular<>'0" & trainicio & "'"
                    End If
                End If
                'cada vez que inicia una pagina
                funciones.Encabezado_Hoja txtcedula.Text, txtapellidos.Text & ", " & txtnombres.Text, txtespecialidad.Text, control, pagina, fila, pag
                fila = fila + 6
                'se seleccionan los trayectos y el regimen PNF, ANUAL o SEMESTRAL
                sql = "select distinct traunidadcurricular as resultado, regimen.desregimen as resultado2 from unidadcurricular, pnf,regimen, pensumunidadcurricular where unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.regimen_codregimen=regimen.codregimen and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
                      " pensumunidadcurricular.pensum_codpensum=" & codpensum & " and pnf.despnf='" & txtespecialidad.Text & condicion & " ORDER BY RESULTADO  ASC"
                'coloca el detalle de las notas y la hoja de resumen de la promocion
                funciones.Cuerpo_Hoja txtcedula.Text, sql, pag, pagina, fila, control, codparticipante, codpromocion, txtapellidos.Text & ", " & txtnombres.Text, txtespecialidad.Text, 1
                rs.MoveNext
            Loop
        End If
    Else                                        'no esta registrado en una promocion
        sql = "select participantecohorte.nivel_codnivel as resultado From participantecohorte,cohorte, participante where participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participantecohorte.participante_codparticipante=" & codparticipante & _
              " and cohorte.pensum_codpensum=" & codpensum & " and participante.usuario_cedusuario='" & txtcedula.Text & "'"
        codnivel = funciones.CampoEntero(sql, cn)
        sql = "select pensumnivel.trapensumnivel as resultado from pensumnivel, participantecohorte,participante where participantecohorte.nivel_codnivel=pensumnivel.nivel_codnivel and participante.codparticipantes=participantecohorte.participante_codparticipante and pensumnivel.pensum_codpensum=" & codpensum & _
              " and pensumnivel.nivel_codnivel=" & codnivel & " and participante.usuario_cedusuario='" & txtcedula.Text & "'"
        trainicio = funciones.CampoEntero(sql, cn)
        sql = "select pensumnivel.finpensumnivel as resultado from pensumnivel, participantecohorte,participante where participantecohorte.nivel_codnivel=pensumnivel.nivel_codnivel and participante.codparticipantes=participantecohorte.participante_codparticipante and pensumnivel.pensum_codpensum=" & codpensum & _
              " and pensumnivel.nivel_codnivel=" & codnivel & " and participante.usuario_cedusuario='" & txtcedula.Text & "'"
        transicion = funciones.CampoEntero(sql, cn)
        If codnivel <> 1 Then
            'trainicio = rs!trapensumnivel
        'Else
            sql = "select pensumnivel.trapensumnivel as resultado from participantecohorte, pensumnivel where participantecohorte.nivel_codnivel=pensumnivel.nivel_codnivel and pensumnivel.pensum_codpensum=" & codpensum & " and pensumnivel.nivel_codnivel=3"
            trainicio = funciones.CampoEntero(sql, cn)
        End If
        'transicion = finpensumnivel
        condicion = "' and cast(traunidadcurricular as int)>=" & trainicio & " and cast(traunidadcurricular as int)<=" & transicion
        If (Mid(txtcohorte.Text, 1, 2) = "IN") And trainicio > 1 Then
                    condicion = condicion & " and traunidadcurricular<>'0" & trainicio & "'"
                Else
                    'aqui se consulta si es egresado de pnf para omitir el trayecto de transicion
                    sql = "select codparticipantecohorte as resultado from participantecohorte, cohorte where participantecohorte.participante_codparticipante=" & codparticipante & " and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and substring(cohorte.descohorte from 1 for 2)='IN'"
                    If (Mid(txtcohorte.Text, 1, 2) = "PS" And funciones.ExisteRegistro(sql, cn)) Then
                        condicion = condicion & " and traunidadcurricular<>'0" & trainicio & "'"
                    End If
                End If
        'cada vez que inicia una pagina
        funciones.Encabezado_Hoja txtcedula.Text, txtapellidos.Text & ", " & txtnombres.Text, txtespecialidad.Text, control, pagina, fila, pag
        fila = fila + 6
        'se seleccionan los trayectos y el regimen PNF, ANUAL o SEMESTRAL
        sql = "select distinct traunidadcurricular as resultado, regimen.desregimen as resultado2 from unidadcurricular, pnf,regimen, pensumunidadcurricular where unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.regimen_codregimen=regimen.codregimen and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
              " pensumunidadcurricular.pensum_codpensum=" & codpensum & " and pnf.despnf='" & txtespecialidad.Text & condicion & " ORDER BY RESULTADO  ASC"
        'coloca el detalle de las notas y la hoja de resumen de la promocion
        funciones.Cuerpo_Hoja txtcedula.Text, sql, pag, pagina, fila, control, codparticipante, codpromocion, txtapellidos.Text & ", " & txtnombres.Text, txtespecialidad.Text, 2
    End If
    funciones.Conectar
    sql = "select apeusuario as resultado from usuario where cedusuario='" & cedusuario & "'"
    If InStr(funciones.CampoString(sql, cn), " ") > 0 Then
        procesadopor = Mid(funciones.CampoString(sql, cn), 1, InStr(funciones.CampoString(sql, cn), " ") - 1)
    Else
        procesadopor = funciones.CampoString(sql, cn)
    End If
    sql = "select nomusuario as resultado from usuario where cedusuario='" & cedusuario & "'"
    If InStr(funciones.CampoString(sql, cn), " ") > 0 Then
        procesadopor = procesadopor & ", " & Mid(funciones.CampoString(sql, cn), 1, InStr(funciones.CampoString(sql, cn), " ") - 1)
    Else
        procesadopor = procesadopor & ", " & funciones.CampoString(sql, cn)
    End If
    cn.Close
    'modifica el pie de pagina
    With AppExcel.Application.ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = "&Procesado por:" & procesadopor
        .CenterFooter = "&""CodabarMedium,Medium""&*" & control & "*"
        .RightFooter = "&Página  &P de &N"
        .PrintHeadings = False
        .PrintGridlines = False
        '.PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        '.Orientation = xlPortrait
        .Draft = False
        '.PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        '.Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 51
    End With
    '/* seguridad * /
    funciones.RegistroEvento funciones.cedusuario, funciones.FormatoFechaConsulta(Date), "Generó Notas Cetificadas de Graduado de " & txtnombres.Text & ", " & txtapellidos.Text & " de " & txtespecialidad.Text & " " & txtpensum.Text & " N° solicitud " & control, funciones.usuario, txtcedula.Text, txtnombres.Text & ", " & txtapellidos.Text
    'cierra la hoja y destruye el objeto de excel
    AppExcel.Application.Visible = True
    AppExcel.Application.ActiveSheet.PageSetup.PrintArea = "$A$1:$J$" & pag - 1
    AppExcel.Application.ActiveWindow.SelectedSheets.PrintPreview
    AppExcel.Application.ActiveWorkbook.Saved = True
    AppExcel.Application.Quit
    FrmConstancias.MousePointer = vbDefault
    MsgBox "Proceso concluido", vbInformation
End Sub



Private Sub Command1_Click()
    MsgBox funciones.fin_del_Mes(Date)
End Sub



Private Sub Form_Load()
    'llena el flexgrid
    FlexNotas.TextMatrix(0, 0) = "Codigo"
    FlexNotas.ColWidth(0) = 800
    FlexNotas.TextMatrix(0, 1) = "Materia"
    FlexNotas.ColWidth(1) = 6000
    FlexNotas.TextMatrix(0, 2) = "Nota"
    FlexNotas.ColWidth(2) = 800
    FlexNotas.TextMatrix(0, 3) = "Sección"
    FlexNotas.ColWidth(2) = 1000
    FlexNotas.TextMatrix(0, 4) = "Creditos"
    FlexNotas.ColWidth(3) = 1000
    FlexNotas.TextMatrix(0, 5) = "Periodo"
    FlexNotas.ColWidth(4) = 1000
    FlexNotas.TextMatrix(0, 6) = "Trayecto"
    FlexNotas.ColWidth(5) = 1000
End Sub


Private Sub txtcedula_KeyPress(keyascii As Integer)
    Dim sql As String
    
    If keyascii = 13 Then
        If Len(txtcedula.Text) > 0 Then
            funciones.Conectar
            sql = "select nomusuario as resultado from usuario, participante where usuario.cedusuario=participante.usuario_cedusuario and cedusuario='" & txtcedula.Text & "'"
            If funciones.ExisteRegistro(sql, cn) Then
                txtnombres = funciones.CampoString(sql, cn)
                sql = "select apeusuario as resultado from usuario, participante where usuario.cedusuario=participante.usuario_cedusuario and cedusuario='" & txtcedula.Text & "'"
                txtapellidos = funciones.CampoString(sql, cn)
                sql = "select count(distinct pnf.despnf) as resultado from participante,participantecohorte, cohorte,pnf where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                        " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and cohorte.pnf_codpnf=pnf.codpnf and participante.usuario_cedusuario='" & txtcedula.Text & "'"
                
                If funciones.CampoEntero(sql, cn) > 1 Then
                    sql = "select pnf.despnf as resultado from participante,participantecohorte, cohorte,pnf where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                        " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and cohorte.pnf_codpnf=pnf.codpnf and participante.usuario_cedusuario='" & txtcedula.Text & "'"
                    funciones.llenarcombobox CboPnf, sql, cn, True
                    CboPnf.Visible = True
                Else
                    sql = "select pnf.despnf as resultado from participante,participantecohorte, cohorte,pnf where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                        " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and cohorte.pnf_codpnf=pnf.codpnf and participante.usuario_cedusuario='" & txtcedula.Text & "'"
                    txtespecialidad.Text = funciones.CampoString(sql, cn)
                    txtespecialidad.SetFocus
                End If
            Else
                MsgBox "el número de cédula no corresponde a ningun participante ", vbCritical
            End If
            cn.Close
        Else
            MsgBox "Debe indicar un número de cédula", vbCritical
        End If
    Else
        
    End If
End Sub

Private Sub txtcohorte_GotFocus()
    Dim sql As String
    Dim totalcreditosnivel As Integer
    Dim codpensum As Integer
    Dim trainicio As Integer
    Dim transicion As Integer
    
    funciones.Conectar
    sql = "select codpensum as resultado from pensum, pnf where pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and pensum.despensum='" & txtpensum.Text & "'"
    codpensum = funciones.CampoEntero(sql, cn)
    sql = "select codcohorte as resultado from cohorte where descohorte='" & txtcohorte.Text & "' and pensum_codpensum=" & codpensum
    codcohorte = funciones.CampoEntero(sql, cn)
    sql = "select trapensumnivel as resultado from cohorte, nivel, pensumnivel where cohorte.nivel_codnivel=nivel.codnivel and nivel.codnivel=pensumnivel.nivel_codnivel and" & _
        " pensumnivel.pensum_codpensum=" & codpensum & " and cohorte.codcohorte=" & codcohorte
    trainicio = funciones.CampoEntero(sql, cn)
    sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & txtcedula.Text & "'"
    codparticipante = funciones.CampoEnteroLargo(sql, cn)
    If trainicio = 1 Then   ' comienza en trayecto inicial
        sql = "select regimen.desregimen as resultado from pnf,regimen where pnf.regimen_codregimen=regimen.codregimen and pnf.despnf='" & txtespecialidad.Text & "'"
        If funciones.ExisteRegistro(sql, cn) = True Then  'regimen de pnf
            sql = "select max(trapensumnivel) as resultado from cohorte,nivel,pensumnivel where cohorte.nivel_codnivel=nivel.codnivel and nivel.codnivel=pensumnivel.nivel_codnivel and" & _
                  " pensumnivel.pensum_codpensum=" & codpensum
            transicion = funciones.CampoEntero(sql, cn)
            condicion1 = " and unidadcurricular.traunidadcurricular<>'0" & transicion & "'"
            sql = "select count(inscripcionucurricular.codinscripcionucurricular) as resultado from inscripcionucurricular, participante,unidadcurricular  where participante.codparticipantes=inscripcionucurricular.participantes_codparticipantes and" & _
                  " inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and unidadcurricular.traunidadcurricular>='3' and calificacion>0 and participante.codparticipantes=" & codparticipante
            If funciones.CampoEntero(sql, cn) > 0 Then  'para el caso que tenga calificaciones despues de la transicion
                condicion1 = " and cast(unidadcurricular.traunidadcurricular as int)>='" & trainicio & "' and unidadcurricular.codunidadcurricular not in (select codunidadcurricular from unidadcurricular where traunidadcurricular='0" & transicion & "')"
                sql = "select trapensum as resultado from pensum where codpensum=" & codpensum
                transicion = funciones.CampoEntero(sql, cn)
                condicion2 = "<='" & transicion & "'"
            Else                                        'para el caso que no tenga calificaciones despues del trayecto inicial
                condicion2 = "<'" & transicion & "'"
            End If
        Else
            condicion1 = " and cast(unidadcurricular.traunidadcurricular as int)>='" & trainicio & "'"
            sql = "select trapensum as resultado from pensum where codpensum=" & codpensum
            transicion = funciones.CampoEntero(sql, cn)
            condicion2 = "<='" & transicion & "'"
        End If
    Else                    'comienza en transicion
        sql = "select trapensum as resultado from pensum where codpensum=" & codpensum
        transicion = funciones.CampoEntero(sql, cn)
        sql = "select count(participantecohorte.codparticipantecohorte) as resultado from participante, participantecohorte,cohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and" & _
            " participante.codparticipantes=" & codparticipante & " and substring(cohorte.descohorte from 1 for 2)<>'PS'"
        If funciones.CampoEntero(sql, cn) > 0 Then
            condicion1 = " and cast(unidadcurricular.traunidadcurricular as int)>=" & trainicio
        Else
            condicion1 = " and cast(unidadcurricular.traunidadcurricular as int)>=" & trainicio & " and unidadcurricular.traunidadcurricular <>'0" & trainicio & "' "
        End If
        condicion2 = "<='" & transicion & "'"
    End If
    
    sql = "select count(distinct  unidadcurricular.codunidadcurricular) AS resultado" & _
        " FROM participante, participantecohorte, cohorte, pensum, pensumunidadcurricular, unidadcurricular where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
        " participantecohorte.cohorte_codcohorte = cohorte.codcohorte and cohorte.pensum_codpensum = pensum.codpensum AND pensum.codpensum = pensumunidadcurricular.pensum_codpensum AND pensumunidadcurricular.unidadcurricular_codunidadcurricular =" & _
        " unidadcurricular.codunidadcurricular AND participante.codparticipantes = " & codparticipante & " AND pensumunidadcurricular.pensum_codpensum=" & codpensum & condicion1 & " and " & _
        " unidadcurricular.traunidadcurricular " & condicion2 '& " order by trayecto Asc"
    
    cantidaducpensum = funciones.CampoEntero(sql, cn)
    'sql = "select pensum.codpensum as resultado from pensum, pnf where pensum.pnf_codpnf=pnf.codpnf and pensum.despensum='" & txtpensum.Text & "' and pnf.despnf='" & txtespecialidad.Text & "'"
    'codpensum = funciones.CampoEntero(sql, cn)
    sql = "select mencion.codmencion as resultado from mencion, pnf,pensum where mencion.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and pensum.despensum='" & txtpensum.Text & "' and pnf.despnf='" & txtespecialidad.Text & "'"
    codmencion = funciones.CampoEntero(sql, cn)
    'codmencion = 14
    'sql = "select codigo, unidadcurricular,creditos, max(calificacion) as calificacion, periodo,seccion,trayecto from reporte_notas, pensumunidadcurricular where reporte_notas.codigo=pensumunidadcurricular.unidadcurricular_codunidadcurricular and pensumunidadcurricular.pensum_codpensum=" & codpensum & " and reporte_notas.cedula='" & txtcedula.Text & "' and condicion='Aprobado' " & _
    '      "group by codigo, unidadcurricular, creditos, periodo, trayecto, seccion order by trayecto, unidadcurricular asc"  'cast(substring(periodo from 3 for 4) as  int), cast(substring(periodo from 1 for 1) as int) asc,unidadcurricular asc
    sql = "select inscripcionucurricular.unidadcurricular_codunidadcurricular as codigo, unidadcurricular.desunidadcurricular as unidadcurricular, inscripcionucurricular.calificacion, secciones.nomsecciones as seccion, unidadcurricular.ucunidadcurricular as creditos, periodosacademicos.desperiodosacademicos as periodo, unidadcurricular.traunidadcurricular as trayecto from inscripcionucurricular,participante,unidadcurricular," & _
          " secciones,periodosacademicos,mencionunidadcurricular where participantes_codparticipantes=participante.codparticipantes and inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and inscripcionucurricular.secciones_codsecciones=secciones.codsecciones and inscripcionucurricular.periodosacademicos_codperiodosacademicos=periodosacademicos.codperiodosacademicos and" & _
          " unidadcurricular.codunidadcurricular=mencionunidadcurricular.unidadcurricular_codunidadcurricular and mencionunidadcurricular.mencion_codmencion=" & codmencion & " and participante.usuario_cedusuario='" & txtcedula.Text & "'"
    'si tiene notas registradas
    If funciones.ExisteRegistro(sql, cn) Then
        Set rs = cn.Execute(sql)
        fila = 1
        Do While Not rs.EOF
            FlexNotas.Rows = fila + 1
            FlexNotas.TextMatrix(fila, 0) = rs!codigo
            FlexNotas.TextMatrix(fila, 1) = rs!unidadcurricular
            FlexNotas.TextMatrix(fila, 2) = rs!calificacion
            FlexNotas.TextMatrix(fila, 3) = rs!seccion
            FlexNotas.TextMatrix(fila, 4) = rs!creditos
            FlexNotas.TextMatrix(fila, 5) = rs!periodo
            totalcreditosnota = totalcreditosnota + rs!creditos * rs!calificacion
            totaluc = totaluc + rs!creditos
            If rs!trayecto = "01" Then
                FlexNotas.TextMatrix(fila, 6) = "Inicial"
            ElseIf rs!trayecto = "03" Or rs!trayecto = "04" Then
                FlexNotas.TextMatrix(fila, 6) = "Transición"
            Else
                FlexNotas.TextMatrix(fila, 6) = rs!trayecto
            End If
            rs.MoveNext
            fila = fila + 1
        Loop
        FlexNotas.Rows = fila + 1
        FlexNotas.TextMatrix(fila, 5) = "I.A.G ="
        FlexNotas.TextMatrix(fila, 6) = Round(totalcreditosnota / totaluc, 2)
        
        'sql = "select pensum.codpensum as resultado from pensum,pnf where pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and pensum.despensum='" & txtpensum.Text & "'"
        'codpensum = funciones.CampoEntero(sql, cn)
        'sql = "select sum(uccreditostrayecto) as resultado from creditostrayecto where cast(tracreditostrayecto as int)>=3 and pensum_codpensum=" & codpensum
        'totalcreditosnivel = funciones.CampoEntero(sql, cn)
        'If totalcreditosnivel <= totaluc Then
            cmdnotascertificadas.Enabled = True
        'Else
        '    MsgBox "al participante le faltan " & totalcreditosnivel - totaluc & " unidades crédito por aprobar", vbCritical
        '    cmdnotascertificadas.Enabled = False
        'End If
    End If
    'llenar la lista de semestre
    CboSemestre.Clear
     For i = CInt(trainicio) To CInt(transicion)
        If i = CInt(trainicio) Then
            If "0" & i = "01" Then
                CboSemestre.AddItem ("Inicial")
            Else
                CboSemestre.AddItem ("Transición")
            End If
            CboSemestre.AddItem (i)
        Else
            CboSemestre.AddItem (i)
        End If
     Next
    cn.Close
End Sub

Private Sub txtespecialidad_GotFocus()
    Dim sql As String
    
    funciones.Conectar
    sql = "select count(pensum.despensum) as resultado from cohorte,pensum, pnf, participantecohorte, participante where" & _
          " cohorte.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and cohorte.codcohorte=participantecohorte.cohorte_codcohorte and" & _
          " participantecohorte.participante_codparticipante=participante.codparticipantes and pnf.despnf='" & txtespecialidad.Text & "' and" & _
          " participante.usuario_cedusuario='" & txtcedula.Text & "'"
    If funciones.CampoEntero(sql, cn) > 1 Then
        sql = "select pensum.despensum as resultado from cohorte,pensum, pnf, participantecohorte, participante where" & _
          " cohorte.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and cohorte.codcohorte=participantecohorte.cohorte_codcohorte and" & _
          " participantecohorte.participante_codparticipante=participante.codparticipantes and pnf.despnf='" & txtespecialidad.Text & "' and" & _
          " participante.usuario_cedusuario='" & txtcedula.Text & "'"
        funciones.llenarcombobox CboPensum, sql, cn, True
        CboPensum.Visible = True
    Else
        sql = "select pensum.despensum as resultado from cohorte,pensum, pnf, participantecohorte, participante where" & _
          " cohorte.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and cohorte.codcohorte=participantecohorte.cohorte_codcohorte and" & _
          " participantecohorte.participante_codparticipante=participante.codparticipantes and pnf.despnf='" & txtespecialidad.Text & "' and" & _
          " participante.usuario_cedusuario='" & txtcedula.Text & "'"
        txtpensum.Text = funciones.CampoString(sql, cn)
        txtpensum.SetFocus
    End If
    cn.Close
End Sub

Private Sub TxtNota_KeyPress(keyascii As Integer)
    Dim fila
    Dim sql As String
    
    If keyascii = 13 Then
        If Len(TxtNota.Text) > 0 Then
            FlexNotas.Rows = FlexNotas.Rows + 1
            fila = FlexNotas.Rows - 1
            funciones.Conectar
            sql = "select unidadcurricular.codunidadcurricular as resultado from unidadcurricular, pensumunidadcurricular,pensum,pnf" & _
                  " where unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
                  " pensumunidadcurricular.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and" & _
                  " unidadcurricular.traunidadcurricular='" & CboSemestre.Text & "' and pensum.despensum='" & txtpensum.Text & "' and unidadcurricular.desunidadcurricular='" & CboMateria.Text & "'"
            FlexNotas.TextMatrix(fila, 0) = funciones.CampoString(sql, cn)
            FlexNotas.TextMatrix(fila, 1) = CboMateria.Text
            FlexNotas.TextMatrix(fila, 2) = TxtNota.Text
            FlexNotas.TextMatrix(fila, 3) = TxtSeccion.Text
            sql = "select unidadcurricular.ucunidadcurricular as resultado from unidadcurricular, pensumunidadcurricular,pensum,pnf" & _
                  " where unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
                  " pensumunidadcurricular.pensum_codpensum=pensum.codpensum and pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and" & _
                  " unidadcurricular.traunidadcurricular='" & CboSemestre.Text & "' and pensum.despensum='" & txtpensum.Text & "' and unidadcurricular.desunidadcurricular='" & CboMateria.Text & "'"
            FlexNotas.TextMatrix(fila, 4) = funciones.CampoEntero(sql, cn)
            FlexNotas.TextMatrix(fila, 5) = CboPeriodo.Text
            FlexNotas.TextMatrix(fila, 6) = CboSemestre.Text
            TxtNota.Text = ""
            TxtSeccion.Text = ""
            cn.Close
            CmdGuardar.Enabled = True
        Else
            MsgBox "Debe indicar un número entre 1 y 20", vbCritical
        End If
    Else
    
    End If
End Sub

Private Sub TxtNotaElectiva_Click()
    If keyascii = 13 Then
        If Len(TxtNotaElectiva.Text) > 0 Then
            
        Else
            MsgBox "Debe indicar un número entre 1 y 20", vbCritical
        End If
    Else
    
    End If
End Sub

Private Sub txtpensum_GotFocus()
    Dim sql As String
    
    cmdnotascertificadas.Enabled = True
    funciones.Conectar
    sql = "select cohorte.descohorte as resultado from participante,participantecohorte, cohorte, pensum,pnf" & _
        " where participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and " & _
        " cohorte.pensum_codpensum=pensum.codpensum and cohorte.pnf_codpnf=pnf.codpnf and pnf.despnf='" & txtespecialidad.Text & "' and pensum.despensum='" & txtpensum.Text & "' and participante.usuario_cedusuario='" & txtcedula.Text & "'"
    txtcohorte.Text = funciones.CampoString(sql, cn)
    txtcohorte.SetFocus
    cn.Close
    
End Sub


Private Sub TxtSeccion_Change()
    TxtSeccion.Text = UCase(TxtSeccion.Text)
    TxtSeccion.SelStart = Len(TxtSeccion.Text)
End Sub

Private Sub TxtSeccion_KeyPress(keyascii As Integer)
    If keyascii = 13 Then
        If Len(TxtSeccion.Text) > 0 Then
            TxtNota.SetFocus
        Else
            MsgBox "Debe indicar una sección", vbCritical
        End If
    End If
End Sub
