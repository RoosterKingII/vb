VERSION 5.00
Begin VB.Form FrmListadoPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera listado de Promoción"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7920
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.OptionButton Option2 
         Caption         =   "Con I.A.G y Puesto"
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Con Libro y Folio"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox CboPromocion 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promoción"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   750
      End
   End
End
Attribute VB_Name = "FrmListadoPromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboPromocion_Click()
    If CboPromocion.Text <> "Seleccione" Then
        cmdgenerar.Enabled = True
    Else
        cmdgenerar.Enabled = False
    End If
End Sub

Private Sub cmdgenerar_Click()
    Dim sql As String
    Dim titulopromo As String
    Dim esp As String
    Dim fecha As String
    
    Dim rs As Recordset
    Dim cont As Integer
    Dim fila As Integer
    Dim pag As Integer
    
    'crea el objeto de excel y abre la aplicacion
    Set AppExcel = CreateObject("Excel.application")
    AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "PLANTILLAS\listado1.xls"  'para abrir el libro
    AppExcel.Application.Windows("listado1.XLS").Activate
    AppExcel.Application.Visible = True
    
    funciones.Conectar
        titulopromo = Mid(CboPromocion.Text, 1, InStr(1, CboPromocion.Text, " EN ") - 1)
        esp = Mid(CboPromocion.Text, InStr(1, CboPromocion.Text, " EN ") + 4, (InStr(CboPromocion.Text, " (") - 1) - (InStr(1, CboPromocion.Text, " EN ") + 3))
        fecha = Mid(CboPromocion.Text, Len(CboPromocion.Text) - 2, 2) & "-"
        fecha = fecha & Mid(CboPromocion.Text, Len(CboPromocion.Text) - 5, 2) & "-"
        fecha = fecha & Mid(CboPromocion.Text, Len(CboPromocion.Text) - 10, 4)
        sql = "select codpromocion as resultado from autenticidad_titulo where promocion='" & titulopromo & "' and titulopromo='" & esp & "' and fecha='" & fecha & "'"
        codpromocion = funciones.CampoEntero(sql, cn)
        If Option2.Value Then
            sql = "select usuario.cedusuario as cedula, usuario.apeusuario as apellidos, usuario.nomusuario as nombres, participantepromocion.iaparticipantepromocion as iag, participantepromocion.pgparticipantepromocion as puesto" & _
                  " From participantepromocion, participante, usuario where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario=usuario.cedusuario and" & _
                  " csparticipantepromocion='TRUE' and ffparticipantepromocion='TRUE' and promocion_codpromocion=" & codpromocion & " order by puesto asc"
            AppExcel.Application.Cells(9, 5).Value = "I.A.G"
            AppExcel.Application.Cells(9, 6).Value = "Puesto"
        Else
            'sql = "select usuario.cedusuario as cedula, usuario.apeusuario as apellidos, usuario.nomusuario as nombres, participantepromocion.liparticipantepromocion as libro, participantepromocion.foparticipantepromocion as folio" & _
            '      " from participantepromocion, participante,usuario where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario=usuario.cedusuario and promocion_codpromocion=" & codpromocion & " order by apellidos asc, nombres asc"
            sql = "select usuario.cedusuario as cedula, usuario.apeusuario as apellidos, usuario.nomusuario as nombres, participantepromocion.liparticipantepromocion as libro, participantepromocion.foparticipantepromocion as folio" & _
                  " From participantepromocion, participante, usuario where participantepromocion.participantes_codparticipantes=participante.codparticipantes and participante.usuario_cedusuario=usuario.cedusuario and" & _
                  "  promocion_codpromocion=" & codpromocion & " order by libro, folio asc"
            AppExcel.Application.Cells(9, 5).Value = "Libro"
            AppExcel.Application.Cells(9, 6).Value = "Folio"
        End If
        Set rs = cn.Execute(sql)
        If Not rs.BOF Then
            AppExcel.Application.Visible = True
            cont = 1
            fila = 11
            pag = 0
            AppExcel.Application.Cells(6, 1).Value = titulopromo & " EN " & esp
            AppExcel.Application.Cells(7, 1).Value = funciones.fechacompleta(fecha)
            Do While Not rs.EOF
                'funciones.Colocar_Centrado CStr(cont), pag + fila, 1, xlCenter
                'funciones.Colocar_Centrado rs!cedula, pag + fila, 2, xlJustify
                'funciones.Colocar_Centrado rs!apellidos, pag + fila, 3, xlJustify
                'funciones.Colocar_Centrado rs!nombres, pag + fila, 4, xlJustify
                'funciones.Colocar_Centrado rs!libro, pag + fila, 5, xlCenter
                'funciones.Colocar_Centrado rs!folio, pag + fila, 6, xlCenter
                AppExcel.Application.Cells(pag + fila, 1).Value = cont
                AppExcel.Application.Cells(pag + fila, 2).Value = rs!cedula
                AppExcel.Application.Cells(pag + fila, 3).Value = rs!apellidos
                AppExcel.Application.Cells(pag + fila, 4).Value = rs!nombres
                If Option2.Value Then
                    If Len(rs!iag) > 2 And Len(rs!iag) < 5 Then
                        AppExcel.Application.Cells(pag + fila, 5).Value = rs!iag & "0"
                    Else
                        AppExcel.Application.Cells(pag + fila, 5).Value = rs!iag
                    End If
                    AppExcel.Application.Cells(pag + fila, 6).Value = rs!puesto
                Else
                    AppExcel.Application.Cells(pag + fila, 5).Value = rs!libro
                    AppExcel.Application.Cells(pag + fila, 6).Value = rs!folio
                End If
                
                rs.MoveNext
                cont = cont + 1
                fila = fila + 1
                If fila = 68 Then
                    
                    pag = pag + fila - 1
                    fila = 11
                    funciones.Encabezado_listado_promocion pag + 1, 1
                End If
            Loop
            If Option2.Value Then
                fila = fila + 2
                If fila = 68 Then
                    
                    pag = pag + fila - 1
                    fila = 11
                    funciones.Encabezado_listado_promocion pag + 1, 1
                End If
                sql = "select indpromocion as resultado from promocion where codpromocion=" & codpromocion
                AppExcel.Application.Cells(pag + fila - 1, 3).Value = "Indice de la Promoción= " & funciones.CampoDouble(sql, cn)
                '/* seguridad * /
                funciones.RegistroEvento funciones.cedusuario, funciones.FormatoFechaConsulta(Date), "Generó listado de promoción con I.A.G y puesto de la " & AppExcel.Application.Cells(6, 1).Value & " del " & AppExcel.Application.Cells(7, 1).Value, funciones.usuario, "", ""
            Else
                '/* seguridad * /
                funciones.RegistroEvento funciones.cedusuario, funciones.FormatoFechaConsulta(Date), "Generó listado de promoción con Libro y Folio de la " & AppExcel.Application.Cells(6, 1).Value & " del " & AppExcel.Application.Cells(7, 1).Value, funciones.usuario, "", ""
            End If
            pag = pag + fila - 1
            AppExcel.Application.ActiveSheet.PageSetup.PrintArea = "$A$1:$F$" & pag
            AppExcel.Application.ActiveWindow.SelectedSheets.PrintPreview
        End If
    cn.Close
    'cierra la hoja y destruye el objeto de excel
    AppExcel.Application.ActiveWorkbook.Saved = True
    AppExcel.Application.Quit
    'Form1.MousePointer = 1
    'MsgBox "Proceso concluido", vbInformation
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sql As String
    
    funciones.Conectar
    sql = "select distinct promocion || ' EN ' ||titulopromo || ' ('|| fecha|| ')' as resultado, titulopromo, fecha from autenticidad_titulo where indicepromo>0 order by fecha desc, titulopromo desc"
    funciones.llenarcombobox CboPromocion, sql, cn, True
    cn.Close
    
End Sub
