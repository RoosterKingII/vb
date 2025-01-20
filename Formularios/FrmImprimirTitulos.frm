VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmImprimirTitulos 
   Caption         =   "Impresión de titulos"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   12735
      Begin VB.Frame Frame2 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3000
         TabIndex        =   12
         Top             =   960
         Width           =   4335
         Begin VB.OptionButton Option2 
            Caption         =   "Todos"
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Especifico"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.ComboBox CboPromocion 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   7815
      End
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txttitulo 
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtespecialidad 
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtfecha 
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid Msgraduandos 
         Height          =   4695
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8281
         _Version        =   393216
         Cols            =   8
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promoción:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Titulo:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   435
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
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   10800
      TabIndex        =   1
      Top             =   6840
      Width           =   2175
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   735
      Left            =   2640
      ScaleHeight     =   675
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   6840
      Width           =   7935
   End
End
Attribute VB_Name = "FrmImprimirTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
        
        sql = "select usuario.cedusuario, usuario.apeusuario, usuario.nomusuario, participantepromocion.iaparticipantepromocion,participantepromocion.csparticipantepromocion,participantepromocion.liparticipantepromocion, participantepromocion.foparticipantepromocion from participantepromocion, participante, usuario " & _
              " where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario = Usuario.cedusuario And" & _
              " participantepromocion.promocion_codpromocion = " & codpromocion & " and usuario.cedusuario='" & cedula & "' and participantepromocion.ffparticipantepromocion=true"
        Set rs = cn.Execute(sql)
        'Command2.Enabled = True
        
        
        If Not rs.BOF Then
            contador = 1
            fil = 1
            cantidad = 0
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
                Msgraduandos.TextMatrix(fil - 1, 2) = rs!apeusuario
                Msgraduandos.TextMatrix(fil - 1, 3) = rs!nomusuario
                Msgraduandos.TextMatrix(fil - 1, 4) = rs!iaparticipantepromocion
                Msgraduandos.TextMatrix(fil - 1, 6) = rs!liparticipantepromocion
                Msgraduandos.TextMatrix(fil - 1, 7) = rs!foparticipantepromocion
                
                
                Msgraduandos.Refresh
                contador = contador + 1
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

        sql = "select usuario.cedusuario, usuario.apeusuario, usuario.nomusuario, participantepromocion.iaparticipantepromocion,participantepromocion.csparticipantepromocion,participantepromocion.liparticipantepromocion, participantepromocion.foparticipantepromocion from participantepromocion, participante, usuario " & _
              " where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario = Usuario.cedusuario And" & _
              " participantepromocion.promocion_codpromocion = " & codpromocion & " and participantepromocion.ffparticipantepromocion=true order by usuario.apeusuario, usuario.nomusuario asc"
        Set rs = cn.Execute(sql)
        'Command2.Enabled = True
        
        
        If Not rs.BOF Then
            contador = 1
            fil = 1
            cantidad = 0
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
                Msgraduandos.TextMatrix(fil - 1, 6) = rs!liparticipantepromocion
                Msgraduandos.TextMatrix(fil - 1, 7) = rs!foparticipantepromocion
                
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

Private Sub CmdImprimir_Click()
    Dim Doc As Word.Application
    Dim i As Integer
    Dim lib As Integer
    Dim fol As Integer
    Dim sql As String
    Dim funcionario As String
    Dim graduando As String
    Dim nombreplantilla As String
    Dim titulo As String
    
    
    Screen.MousePointer = vbHourglass
    
    ProgressBar1.Visible = True
    funciones.Conectar
    titulo = Mid(CboPromocion, InStr(1, CboPromocion.Text, " EN ") + 4, Len(CboPromocion.Text))
    If InStr(titulo, "(") > 0 Then
        titulo = Mid(titulo, 1, InStr(titulo, "(") - 2)
    End If
    If txttitulo = "TECNICO SUPERIOR UNIVERSITARIO" Then
        If Len(titulo) <= 12 Then
            nombreplantilla = "PlantillaTituloTsu2.Doc"
        Else
            nombreplantilla = "PlantillaTituloTsu.Doc"
        End If
    Else
        nombreplantilla = "PlantillaTituloIng.Doc"
    End If
   
    For i = 1 To Msgraduandos.Rows - 1
        lib = Msgraduandos.TextMatrix(i, 6)
        fol = Msgraduandos.TextMatrix(i, 7)
        graduando = UCase(Msgraduandos.TextMatrix(i, 3)) & " " & UCase(Msgraduandos.TextMatrix(i, 2))
        If lib <> "000" Then
            Set Doc = CreateObject("Word.application")
            Doc.Application.Documents.Open FileName:=funciones.DirectorioActual & "\Plantillas\" & nombreplantilla    'para abrir el documento
            Doc.Application.Documents(nombreplantilla).Activate
            Doc.Application.Visible = True
        
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="libro"
            Doc.Selection.TypeText Text:=lib
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="folio"
            Doc.Selection.TypeText Text:=funciones.formatofolio(fol)
            
            If Len(Msgraduandos.TextMatrix(i, 1)) = 7 Then
                cedula = Mid(Msgraduandos.TextMatrix(i, 1), 1, 1) & "." & Mid(Msgraduandos.TextMatrix(i, 1), 2, 3) & "." & Mid(Msgraduandos.TextMatrix(i, 1), 5, 3)
            Else
                cedula = Mid(Msgraduandos.TextMatrix(i, 1), 1, 2) & "." & Mid(Msgraduandos.TextMatrix(i, 1), 3, 3) & "." & Mid(Msgraduandos.TextMatrix(i, 1), 6, 3)
            End If
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="cedula"
            Doc.Selection.TypeText Text:=cedula
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="graduando"
            Doc.Selection.TypeText Text:=StrConv(graduando, vbProperCase)
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="graduando2"
            Doc.Selection.TypeText Text:=StrConv(graduando, vbProperCase)

            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="titulo3"
            Doc.Selection.TypeText Text:=StrConv(titulo, vbProperCase)
          
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="titulo1"
            Doc.Selection.TypeText Text:=StrConv(titulo, vbProperCase)
            
            Doc.Selection.GoTo What:=wdGoToBookmark, Name:="titulo2"
            Doc.Selection.TypeText Text:=StrConv(titulo, vbProperCase)
            
            'If Len(Msgraduandos.TextMatrix(i, 5)) > 5 Then
            '    titulo = titulo & " Mención " & StrConv(Msgraduandos.TextMatrix(i, 5), vbProperCase)
            'End If
            Doc.ActiveDocument.PrintOut Item:=wdPrintDocumentContent, Copies:=1, Pages:="1"
            sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & Msgraduandos.TextMatrix(i, 1) & "'"
            codparticipante = funciones.CampoEnteroLargo(sql, cn)
            'sql = "update participantepromocion set liparticipantepromocion=" & lib & ", foparticipantepromocion=" & funciones.formatofolio(fol) & " where participantes_codparticipantes=" & codparticipante
            'cn.Execute (sql)
            graduando = Msgraduandos.TextMatrix(i, 2) & " " & Msgraduandos.TextMatrix(i, 3)
            sql = "select nomapeusuario as resultado from usuario where cedusuario='" & funciones.cedusuario & "'"
            funcionario = funciones.CampoString(sql, cn)
            funciones.RegistroEvento funciones.cedusuario, Date, "Impresión de Titulo de  " & txttitulo.Text & " en " & titulo & " Folio " & funciones.formatofolio(fol) & " Libro " & lib, funcionario, Msgraduandos.TextMatrix(i, 1), graduando
            Doc.ActiveDocument.Saved = True
            Doc.Application.Quit
            Set Doc = Nothing
        End If
        'ProgressBar1.Value = (i / (Msgraduandos.Rows - 1)) * 100
     Next i
    
    Screen.MousePointer = vbNormal
    MsgBox "Proceso concluido", vbInformation
    ProgressBar1.Visible = False
    'Command2.Enabled = False
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
End Sub
