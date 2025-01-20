VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmGenerarFolios 
   Caption         =   "Generación de Folios de Grado"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form7"
   ScaleHeight     =   8085
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ProgressBar1 
      Height          =   735
      Left            =   2640
      ScaleHeight     =   675
      ScaleWidth      =   7875
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   1200
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "Especifico"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Todos"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir folios"
      Height          =   735
      Left            =   10800
      TabIndex        =   6
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insertar en promocion"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.TextBox txtfecha 
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtespecialidad 
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txttitulo 
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   8040
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid Msgraduandos 
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   8
      End
      Begin VB.ComboBox CboPromocion 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Especialidad"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Titulo:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promoción:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmGenerarFolios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cantidad As Integer 'cantidad de folios que se procesaran

Function Porcentaje_aprobado_Pensum(cedula As String) As Double
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim totalucpensum As Integer
    Dim totalucaprob As Integer
    Dim codpnf As Integer
    Dim codcohorte As Integer
    Dim despromo As String
    Dim titpromo As String
    
    despromo = Mid(CboPromocion.Text, 1, InStr(CboPromocion.Text, " EN ") - 1)
    titpromo = Mid(CboPromocion.Text, InStr(CboPromocion.Text, " EN ") + 4, Len(CboPromocion.Text))
    sql = "select codpromocion as resultado from promocion where despromocion='" & despromo & "' and titpromocion='" & titpromo & "'"
    codpromocion = funciones.CampoEntero(sql, cn)
    sql = "select promocion.nivel_codnivel as resultado from cohortepromocion, promocion where cohortepromocion.promocion_codpromocion=promocion.codpromocion and" & _
        " promocion.codpromocion=" & codpromocion
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        Do While Not rs.EOF
            
            rs.MoveNext
        Loop
    End If
    sql = "select cohorte.pnf_codpnf as resultado from participante, participantecohorte,cohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
        " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participantecohorte.actparticipantecohorte='TRUE' and participante.usuario_cedusuario='" & cedula & "'"
    codpnf = funciones.CampoEntero(sql, cn)
    
    sql = "select cohorte.codcohorte as resultado from participante, participantecohorte,cohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
        " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participantecohorte.actparticipantecohorte='TRUE' and participante.usuario_cedusuario='" & cedula & "'"
    codcohorte = funciones.CampoEntero(sql, cn)
    
    sql = "select count(cohorte_codcohorte"
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        
    Porcentaje_aprobado_Pensum = rs!resultado
    End If
End Function
Private Sub CboPromocion_Click()
    If CboPromocion.Text <> "Seleccione" Then
        If Len(CboPromocion.Text) > 0 Then
            Frame2.Enabled = True
            cmdbuscar.Enabled = True
        End If
    Else
        cmdbuscar.Enabled = False
    End If
End Sub

Private Sub cmdbuscar_Click()
    Dim cedula As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim contador As Integer
    Dim fil As Integer
    Dim condicion1 As String
    Dim condicion2 As String
    Dim codpensum As Integer
    Dim codpromocion As Integer
    Dim nivelpromocion As Integer
    Dim libro As Integer
    Dim folio As Integer
    
    If Option1.Value Then
        cedula = InputBox("Cédula Graduando:")
        funciones.Conectar
        condicion1 = Mid(CboPromocion.Text, 1, InStr(CboPromocion.Text, " EN ") - 1)
        condicion2 = Mid(CboPromocion.Text, InStr(CboPromocion.Text, " EN ") + 4, Len(CboPromocion.Text))
        sql = "select codpromocion as resultado From promocion where promocion.despromocion ='" & condicion1 & "' AND promocion.titpromocion='" & condicion2 & "'"
        codpromocion = funciones.CampoEntero(sql, cn)
        sql = "select titulos_codtitulos as resultado From promocion where promocion.despromocion ='" & condicion1 & "' AND promocion.titpromocion='" & condicion2 & "'"
        codtitulo = funciones.CampoEntero(sql, cn)
        sql = "select destitulos as resultado from titulos where codtitulos=" & codtitulo
        txttitulo.Text = funciones.CampoString(sql, cn)
        txtespecialidad.Text = condicion2
        sql = "select fecpromocion as resultado from promocion where codpromocion=" & codpromocion
        txtfecha.Text = funciones.fechacompleta(funciones.CampoString(sql, cn))
        sql = "select nivel_codnivel as resultado from promocion where codpromocion=" & codpromocion
        nivel = funciones.CampoEntero(sql, cn)
        
        If nivel = 1 Then
            condicion1 = "=" & nivel
        Else
            condicion1 = ">=" & nivel
        End If
        sql = "select max(cast(participantepromocion.liparticipantepromocion as int)) as resultado from participantepromocion, promocion where promocion.codpromocion=participantepromocion.promocion_codpromocion and" & _
              " promocion.nivel_codnivel" & condicion1
        libro = funciones.CampoEntero(sql, cn)
        'libro = 17
        sql = "select max(cast(participantepromocion.foparticipantepromocion as int)) as resultado from participantepromocion, promocion where promocion.codpromocion=participantepromocion.promocion_codpromocion and" & _
              " promocion.nivel_codnivel" & condicion1 & " and participantepromocion.liparticipantepromocion='" & libro & "'"
        folio = funciones.CampoEntero(sql, cn) + 1
        sql = "select usuario.cedusuario, usuario.apeusuario, usuario.nomusuario, participantepromocion.iaparticipantepromocion,participantepromocion.csparticipantepromocion,participantepromocion.liparticipantepromocion, participantepromocion.foparticipantepromocion from participantepromocion, participante, usuario " & _
              " where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario = Usuario.cedusuario And" & _
              " participantepromocion.promocion_codpromocion = " & codpromocion & " and usuario.cedusuario='" & cedula & "'"
        Set rs = cn.Execute(sql)
        'Command2.Enabled = True
        
        
        If Not rs.BOF Then
            contador = 1
            fil = 1
            cantidad = 0
            'Msgraduandos.SelectionMode = flexSelectionByRow
            Do While Not rs.EOF
                fil = fil + 1
                Msgraduandos.Rows = fil
                sql = "select mencion.desmencion as resultado from mencion,participantecohorte, participante,participantepromocion where participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.mencion_codmencion =mencion.codmencion and" & _
                    " participante.usuario_cedusuario='" & rs!cedusuario & "' and participantepromocion.participantecohorte_codparticipantecohorte= participantecohorte.codparticipantecohorte and participantepromocion.promocion_codpromocion=" & codpromocion
                mencion = funciones.CampoString(sql, cn)
                If mencion <> "000" Then
                    Msgraduandos.TextMatrix(fil - 1, 5) = mencion
                End If
                
                
               
                Msgraduandos.TextMatrix(fil - 1, 0) = contador
                Msgraduandos.TextMatrix(fil - 1, 1) = rs!cedusuario
                'Msgraduandos.TextMatrix(fil - 1, 1).ForeColor = vbYellow
                Msgraduandos.TextMatrix(fil - 1, 2) = rs!apeusuario
                Msgraduandos.TextMatrix(fil - 1, 3) = rs!nomusuario
                Msgraduandos.TextMatrix(fil - 1, 4) = rs!iaparticipantepromocion
                If rs!liparticipantepromocion = "0" Then
                    If rs!csparticipantepromocion = True Then
                        Msgraduandos.TextMatrix(fil - 1, 6) = libro
                        cantidad = cantidad + 1
                    Else
                        Msgraduandos.TextMatrix(fil - 1, 6) = "000"
                    End If
                    If rs!csparticipantepromocion = True Then
                        Msgraduandos.TextMatrix(fil - 1, 7) = folio
                    Else
                        Msgraduandos.TextMatrix(fil - 1, 7) = "000"
                    End If
                Else
                    'Command2.Enabled = False
                    Msgraduandos.TextMatrix(fil - 1, 6) = rs!liparticipantepromocion
                    Msgraduandos.TextMatrix(fil - 1, 7) = rs!foparticipantepromocion
                End If
                Msgraduandos.Refresh
                contador = contador + 1
                If rs!csparticipantepromocion = True Then
                    folio = folio + 1
                End If
                If folio > 500 Then
                    libro = libro + 1
                    folio = 1
                End If
                rs.MoveNext
            Loop
        Else
            MsgBox "no hay graduandos asociados a la promoción", vbCritical
        End If
        'cn.Close
    Else
        funciones.Conectar
        condicion1 = Mid(CboPromocion.Text, 1, InStr(CboPromocion.Text, " EN ") - 1)
        condicion2 = Mid(CboPromocion.Text, InStr(CboPromocion.Text, " EN ") + 4, Len(CboPromocion.Text))
        sql = "select codpromocion as resultado From promocion where promocion.despromocion ='" & condicion1 & "' AND promocion.titpromocion='" & condicion2 & "'"
        codpromocion = funciones.CampoEntero(sql, cn)
        sql = "select titulos_codtitulos as resultado From promocion where promocion.despromocion ='" & condicion1 & "' AND promocion.titpromocion='" & condicion2 & "'"
        codtitulo = funciones.CampoEntero(sql, cn)
        sql = "select destitulos as resultado from titulos where codtitulos=" & codtitulo
        txttitulo.Text = funciones.CampoString(sql, cn)
        txtespecialidad.Text = condicion2
        sql = "select fecpromocion as resultado from promocion where codpromocion=" & codpromocion
        txtfecha.Text = funciones.fechacompleta(funciones.CampoString(sql, cn))
        sql = "select nivel_codnivel as resultado from promocion where codpromocion=" & codpromocion
        nivel = funciones.CampoEntero(sql, cn)
        
        If nivel = 1 Then
            condicion1 = "=" & nivel
        Else
            condicion1 = ">=" & nivel
        End If
        sql = "select max(cast(participantepromocion.liparticipantepromocion as int)) as resultado from participantepromocion, promocion where promocion.codpromocion=participantepromocion.promocion_codpromocion and" & _
              " promocion.nivel_codnivel" & condicion1
        libro = funciones.CampoEntero(sql, cn)
        'libro = 17
        sql = "select max(cast(participantepromocion.foparticipantepromocion as int)) as resultado from participantepromocion, promocion where promocion.codpromocion=participantepromocion.promocion_codpromocion and" & _
              " promocion.nivel_codnivel" & condicion1 & " and participantepromocion.liparticipantepromocion='" & libro & "'"
        folio = funciones.CampoEntero(sql, cn) + 1
        sql = "select usuario.cedusuario, usuario.apeusuario, usuario.nomusuario, participantepromocion.iaparticipantepromocion,participantepromocion.csparticipantepromocion,participantepromocion.liparticipantepromocion, participantepromocion.foparticipantepromocion from participantepromocion, participante, usuario " & _
              " where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario = Usuario.cedusuario And" & _
              " participantepromocion.promocion_codpromocion = " & codpromocion & " order by usuario.apeusuario, usuario.nomusuario asc"
        Set rs = cn.Execute(sql)
        'Command2.Enabled = True
        
        
        If Not rs.BOF Then
            contador = 1
            fil = 1
            cantidad = 0
            'Msgraduandos.SelectionMode = flexSelectionByRow
            Do While Not rs.EOF
                fil = fil + 1
                Msgraduandos.Rows = fil
                sql = "select mencion.desmencion as resultado from mencion,participantecohorte, participante,participantepromocion where participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.mencion_codmencion =mencion.codmencion and" & _
                    " participante.usuario_cedusuario='" & rs!cedusuario & "' and participantepromocion.participantecohorte_codparticipantecohorte= participantecohorte.codparticipantecohorte and participantepromocion.promocion_codpromocion=" & codpromocion
                mencion = funciones.CampoString(sql, cn)
                If mencion <> "000" Then
                    Msgraduandos.TextMatrix(fil - 1, 5) = mencion
                End If
                
                
               
                Msgraduandos.TextMatrix(fil - 1, 0) = contador
                Msgraduandos.TextMatrix(fil - 1, 1) = rs!cedusuario
                'Msgraduandos.TextMatrix(fil - 1, 1).ForeColor = vbYellow
                Msgraduandos.TextMatrix(fil - 1, 2) = rs!apeusuario
                Msgraduandos.TextMatrix(fil - 1, 3) = rs!nomusuario
                Msgraduandos.TextMatrix(fil - 1, 4) = rs!iaparticipantepromocion
                If rs!liparticipantepromocion = "0" Then
                    If rs!csparticipantepromocion = True Then
                        Msgraduandos.TextMatrix(fil - 1, 6) = libro
                        cantidad = cantidad + 1
                    Else
                        Msgraduandos.TextMatrix(fil - 1, 6) = "000"
                    End If
                    If rs!csparticipantepromocion = True Then
                        Msgraduandos.TextMatrix(fil - 1, 7) = folio
                    Else
                        Msgraduandos.TextMatrix(fil - 1, 7) = "000"
                    End If
                Else
                    'Command2.Enabled = False
                    Msgraduandos.TextMatrix(fil - 1, 6) = rs!liparticipantepromocion
                    Msgraduandos.TextMatrix(fil - 1, 7) = rs!foparticipantepromocion
                End If
                Msgraduandos.Refresh
                contador = contador + 1
                If rs!csparticipantepromocion = True Then
                    folio = folio + 1
                End If
                If folio > 500 Then
                    libro = libro + 1
                    folio = 1
                End If
                rs.MoveNext
            Loop
        Else
            MsgBox "no hay graduandos asociados a la promoción", vbCritical
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim iag As Double
    Dim sql As String
    Dim codpensum As Integer
    
    funciones.Conectar
        condicion1 = Mid(CboPromocion.Text, 1, InStr(CboPromocion.Text, " EN ") - 1)
        condicion2 = Mid(CboPromocion.Text, InStr(CboPromocion.Text, " EN ") + 4, Len(CboPromocion.Text))
        sql = "select cohortepromocion.cohorte_codcohorte as resultado1,promocion.nivel_codnivel as resultado2, promocion.codpromocion as resultado3 From promocion, cohortepromocion where promocion.codpromocion=cohortepromocion.promocion_codpromocion and" & _
              " promocion.despromocion ='" & condicion1 & "' AND promocion.titpromocion='" & condicion2 & "'"
        Set rs = cn.Execute(sql)
        If Not rs.BOF Then
            contador = 1
            fil = 1
            Do While Not rs.EOF
                'sql = "select usuario.cedusuario as cedula,usuario.nomusuario as nombre, usuario.apeusuario as apellido,cohorte.descohorte, cohorte.pnf_codpnf, cohorte.pensum_codpensum from solvencia, participante, usuario, participantecohorte, cohorte where usuario.cedusuario=participante.usuario_cedusuario and" & _
                '      " participante.codparticipantes=participantecohorte.participante_codparticipante and participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participante.codparticipantes=solvencia.participante_codparticipante and consolvencia ='TRUE' and participantecohorte.actparticipantecohorte='true' and " & _
                '      " cohorte.codcohorte=" & rs!resultado1 & " and participantecohorte.titulos_codtitulos=2 order by usuario.apeusuario, usuario.nomusuario asc"
                sql = "select participante.usuario_cedusuario as cedula from participante, participantepromocion where participante.codparticipantes=participantepromocion.participantes_codparticipantes and participantepromocion.promocion_codpromocion=" & rs!resultado3
                Set rs2 = cn.Execute(sql)
                If Not rs2.BOF Then
                    sql = "select participantecohorte.cohorte_codcohorte as resultado from participantecohorte, participante where participante.codparticipantes=participantecohorte.participante_codparticipante and participante.usuario_cedusuario='" & rs2!cedula & "'"
                    codcohorte = funciones.CampoEntero(sql, cn)
                    sql = "select pensum_codpensum as resultado  from cohorte where codcohorte=" & codcohorte
                    codpensum = funciones.CampoEntero(sql, cn)
                    codpromocion = rs!resultado3
                    Do While Not rs2.EOF
                        iag = Format(CalcularIAG(rs2!cedula, codpensum), "#0.00")
                        If iag < 12 Then
                            iag = 12
                        End If
                        sql = "select max(codparticipantepromocion) as resultado from participantepromocion"
                        codparticipantepromocion = funciones.proximocodigoregistro(sql, cn)
                        fecha = Year(Date) & "-" & Month(Date) & "-" & Day(Date)
                        sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & rs2!cedula & "'"
                        codparticipante = funciones.CampoEnteroLargo(sql, cn)
                        'sql = "insert into participantepromocion (codparticipantepromocion, participantes_codparticipantes, iaparticipantepromocion, frparticipantepromocion, promocion_codpromocion)" & _
                        '    " values (" & codparticipantepromocion & "," & codparticipante & "," & Replace(CStr(iag), ",", ".") & ",'" & fecha & "'," & rs!resultado3 & ")"
                        sql = "update participantepromocion set iaparticipantepromocion=" & Replace(CStr(iag), ",", ".") & " where participantes_codparticipantes=" & codparticipante
                        cn.Execute (sql)
                        If rs!resultado2 = 2 Then
                           sql = "update participantecohorte set estatus_codestatus=7, titulos_codtitulos=3 where participante_codparticipante=" & codparticipante & " and cohorte_codcohorte=" & rs!resultado1
                        Else
                            sql = "update participantecohorte set titulos_codtitulos=2 where participante_codparticipante=" & codparticipante & " and cohorte_codcohorte=" & rs!resultado1
                        End If
                        cn.Execute (sql)
                        rs2.MoveNext
                    Loop
               End If
                rs.MoveNext
            Loop
            'codpromocion = 6
            sql = "select participantepromocion.codparticipantepromocion,usuario.apeusuario, usuario.nomusuario, participantepromocion.iaparticipantepromocion  from participantepromocion, participante, usuario where " & _
                  " participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario=usuario.cedusuario and" & _
                  " promocion_codpromocion=" & codpromocion & " order by iaparticipantepromocion desc, usuario.apeusuario, usuario.nomusuario asc"
            Set rs = cn.Execute(sql)
            If Not rs.BOF Then
                contador = 1
                Do While Not rs.EOF
                    sql = "update participantepromocion set pgparticipantepromocion=" & contador & "where codparticipantepromocion=" & rs!codparticipantepromocion
                    cn.Execute (sql)
                    contador = contador + 1
                    rs.MoveNext
                Loop
            End If
            MsgBox "Proceso Concluido", vbInformation
        Else
        
        End If
End Sub

Private Sub Command2_Click()
    Dim Doc As Word.Application
    Dim i As Integer
    Dim lib As Integer
    Dim fol As Integer
    Dim sql As String
    Dim funcionario As String
    Dim graduando As String
    
    Screen.MousePointer = vbHourglass
    'lib = InputBox("Indique libro de inicio")
    'fol = InputBox("Indique folio de inicio")
    'ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    funciones.Conectar
    For i = 1 To Msgraduandos.Rows - 1
        lib = Msgraduandos.TextMatrix(i, 6)
        fol = Msgraduandos.TextMatrix(i, 7)
        If lib <> "000" Then
            Set Doc = CreateObject("Word.application")
            Doc.Application.Documents.Open FileName:=funciones.DirectorioActual & "\Plantillas\PlantillaFolio.doc"   'para abrir el documento
            Doc.Application.Documents("PlantillaFolio.doc").Activate
            'Doc.Application.Visible = True
        
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="libro"
            Doc.Selection.TypeText Text:=lib
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="folio"
            Doc.Selection.TypeText Text:=funciones.formatofolio(fol)
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="cedula"
            Doc.Selection.TypeText Text:=Msgraduandos.TextMatrix(i, 1)
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="apellidos"
            Doc.Selection.TypeText Text:=UCase(Msgraduandos.TextMatrix(i, 2))
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="nombres"
            Doc.Selection.TypeText Text:=UCase(Msgraduandos.TextMatrix(i, 3))
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="titulo"
            Doc.Selection.TypeText Text:=txttitulo.Text
            
            If Len(Msgraduandos.TextMatrix(i, 5)) > 5 Then
                Doc.Selection.GoTo What:=wdGoToBookmark, Name:="especialidad"
                Doc.Selection.TypeText Text:=txtespecialidad.Text & " MENCION: " & Msgraduandos.TextMatrix(i, 5)
            Else
                If InStr(txtespecialidad.Text, "(") > 0 Then
                    Doc.Selection.GoTo What:=wdGoToBookmark, Name:="especialidad"
                    Doc.Selection.TypeText Text:=Mid(txtespecialidad.Text, 1, InStr(txtespecialidad.Text, " (") - 1)
                Else
                    Doc.Selection.GoTo What:=wdGoToBookmark, Name:="especialidad"
                    Doc.Selection.TypeText Text:=txtespecialidad.Text
                End If
                'aa = Mid(txtespecialidad.Text, 1, InStr(txtespecialidad.Text, " (") - 1)
            End If
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="fechagrado"
            Doc.Selection.TypeText Text:=txtfecha.Text
            
            'Doc.Selection.GoTo What:=wdGoToBookmark, Name:="lugar1"
            'Doc.Selection.TypeText Text:=UCase(Msgraduandos.TextMatrix(i, 5))
            
            'Doc.Selection.GoTo What:=wdGoToBookmark, Name:="lugar2"
            'Doc.Selection.TypeText Text:=UCase(Msgraduandos.TextMatrix(i, 6))
            
            'Doc.Selection.GoTo What:=wdGoToBookmark, Name:="fechanac"
            'Doc.Selection.TypeText Text:=Msgraduandos.TextMatrix(i, 7)
            Doc.Application.Visible = True
            
            
            Doc.ActiveDocument.PrintOut Item:=wdPrintDocumentContent, Copies:=2, Pages:="1"
            sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & Msgraduandos.TextMatrix(i, 1) & "'"
            codparticipante = funciones.CampoEnteroLargo(sql, cn)
            sql = "update participantepromocion set liparticipantepromocion=" & lib & ", foparticipantepromocion=" & funciones.formatofolio(fol) & " where participantes_codparticipantes=" & codparticipante
            cn.Execute (sql)
            graduando = Msgraduandos.TextMatrix(i, 2) & " " & Msgraduandos.TextMatrix(i, 3)
            sql = "select nomapeusuario as resultado from usuario where cedusuario='" & funciones.cedusuario & "'"
            funcionario = funciones.CampoString(sql, cn)
            funciones.RegistroEvento funciones.cedusuario, Date, "Impresión del Folio " & funciones.formatofolio(fol) & " del libro " & lib, funcionario, Msgraduandos.TextMatrix(i, 1), graduando
            Doc.ActiveDocument.Saved = True
            Doc.Application.Quit
            Set Doc = Nothing
        End If
        'ProgressBar1.Value = (i / cantidad) * 100
     Next i
    
    Screen.MousePointer = vbNormal
    MsgBox "Proceso concluido", vbInformation
    ProgressBar1.Visible = False
    Command2.Enabled = False
End Sub

Private Sub Form_Load()
    Dim sql As String
    funciones.Conectar
    
        sql = "select (despromocion ||' EN '||titpromocion) as resultado from promocion where estpromocion='TRUE' order by codpromocion asc"
        funciones.llenarcombobox CboPromocion, sql, cn, True
    cn.Close
    Msgraduandos.TextMatrix(0, 0) = "Nro"
    Msgraduandos.ColWidth(0) = 500
    Msgraduandos.TextMatrix(0, 1) = "Cédula"
    Msgraduandos.ColWidth(1) = 800
    Msgraduandos.TextMatrix(0, 2) = "Apellidos"
    Msgraduandos.ColWidth(2) = 3000
    Msgraduandos.TextMatrix(0, 3) = "Nombres"
    Msgraduandos.ColWidth(3) = 3000
    Msgraduandos.TextMatrix(0, 4) = "IAG"
    Msgraduandos.ColWidth(4) = 600
    Msgraduandos.TextMatrix(0, 5) = "MENCIÓN"
    Msgraduandos.ColWidth(5) = 3000
    Msgraduandos.TextMatrix(0, 6) = "LIB."
    Msgraduandos.ColWidth(6) = 600
    Msgraduandos.TextMatrix(0, 7) = "FOL."
    Msgraduandos.ColWidth(7) = 600
    'Msgraduandos.TextMatrix(0, 8) = "MENCION"
    'Msgraduandos.ColWidth(3) = 2800
    
    
End Sub


