VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReparar1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "Command5"
      Height          =   615
      Left            =   2160
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Reparar Inscripcion "
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Insertar Inscripcion INDIVIDUAL"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Insertar Inscripcion"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Insertar notas"
      Height          =   615
      Left            =   4080
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Insertar participantes"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Insertar Personal ISR"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Resumen nomina"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "insertar unidad curricular"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insertar Familiar"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insertar periodos academicos"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insertar usuarios"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "FrmReparar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    Dim codparticipante As Long
    
    cont = 0
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "electricidad T02 1-2014.xls"  'para abrir el libro
        AppExcel.Application.Windows("electricidad T02 1-2014.xls").Activate
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            cedula = AppExcel.Application.Cells(fila, 1).Value
            PAIS1 = "000"
            estcivil = 1
            nombre = AppExcel.Application.Cells(fila, 3).Value
            apellido = AppExcel.Application.Cells(fila, 2).Value
            correo = ""
            telefono = ""
            celular = ""
            fechanac = "1900-01-01"
            av = "" 'AppExcel.application.cells(fila, 10).Value
            p1 = "300001" 'AppExcel.application.cells(fila, 11).Value
            m1 = "3001" 'AppExcel.application.cells(fila, 12).Value
            m2 = "3001" 'AppExcel.application.cells(fila, 13).Value
            p2 = "300001" 'AppExcel.application.cells(fila, 14).Value
            e1 = "30" 'AppExcel.application.cells(fila, 15).Value
            e2 = "30" 'AppExcel.application.cells(fila, 16).Value
            inst = 0 'AppExcel.application.cells(fila, 17).Value
            sex = AppExcel.Application.Cells(fila, 7).Value
            'If sex = "M" Then
            '    sex = "1"
            'Else
            '    sex = "0"
            'End If
            paisnac = "000" 'AppExcel.application.cells(fila, 19).Value
            urb = "000" 'AppExcel.application.cells(fila, 20).Value
            casa = "000" ' AppExcel.application.cells(fila, 21).Value
            completo = nombre & " " & apellido 'AppExcel.application.cells(fila, 22).Value
            etnia = "false" 'AppExcel.application.cells(fila, 23).Value
            Code = 0 'AppExcel.application.cells(fila, 24).Value
            disc = "false" 'AppExcel.application.cells(fila, 25).Value
            descri = "N/A" 'AppExcel.application.cells(fila, 26).Value
            Doc = "C" 'AppExcel.application.cells(fila, 27).Value
            resid = 0 'AppExcel.application.cells(fila, 28).Value
            
            Set rs = cn.Execute("select nomusuario from usuario where cedusuario='" & cedula & "'")
            If rs.BOF Then
                
                cn.Execute ("insert into usuario VALUES('" & cedula & "','" & PAIS1 & "'," & estcivil & ",'" & nombre & "','" & apellido & "','" & _
                correo & "','" & telefono & "','" & celular & "','" & fechanac & "','" & av & "','" & p1 & "','" & m1 & "','" & m2 & "','" & p2 & "','" & _
                e1 & "','" & e2 & "'," & inst & ",'" & sex & "','" & paisnac & "','" & urb & "','" & casa & "','" & completo & "','" & etnia & "'," & _
                Code & ",'" & disc & "','" & descri & "','" & Doc & "'," & resid & ")")
                'inserta el usuario como participante
                Set rs = cn.Execute("select max(codparticipantes) as resultado from participante")
                If Not rs.BOF Then
                    codparticipante = rs!resultado + 1
                End If
                cn.Execute ("INSERT INTO PARTICIPANTE VALUES ('" & cedula & "',110,1,11140,1,1," & codparticipante & ")")
            Else
                'inserta el usuario como participante
                Set rs = cn.Execute("select max(codparticipantes) as resultado from participante")
                If Not rs.BOF Then
                    codparticipante = rs!resultado + 1
                End If
                cn.Execute ("INSERT INTO PARTICIPANTE VALUES ('" & cedula & "',110,1,11140,1,1," & codparticipante & ")")
                AppExcel.Application.Range("A" & fila).Select
                With AppExcel.Application.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 60535
                    
                   
                End With
                cont = cont + 1
             End If
            fila = fila + 1
        Loop
        MsgBox "Listo Errores-->" & cont
    'End If
        
End Sub

Private Sub Command10_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    Dim codseccion As Integer
    Dim codparticipante As Long
    Dim codinscripcion As Long
    Dim cant_insc As Integer
    
    cont = 0
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "migracion 6to semestre.xls"  'para abrir el libro
        AppExcel.Application.Windows("migracion 6to semestre.xls").Activate
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
        fila = 1
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            cedula = AppExcel.Application.Cells(fila, 1).Value  'cedula
            PAIS1 = AppExcel.Application.Cells(fila, 2).Value   'unidad curricular
            estcivil = AppExcel.Application.Cells(fila, 3).Value    'seccion
            'codigo de participante
             Set rs = cn.Execute("select codparticipantes as resultado from participante where usuario_cedusuario='" & cedula & "'")
            If Not rs.BOF Then
                codparticipante = rs!resultado
            End If
            'codigo de inscripcion
             Set rs = cn.Execute("select max(codinscripcionucurricular) as resultado from inscripcionucurricular")
            If Not rs.BOF Then
                codinscripcion = rs!resultado + 1
            End If
            'codigo de seccion
            Set rs = cn.Execute("select codsecciones as resultado from secciones where unidadcurricular_codunidadcurricular='" & PAIS1 & "' and nomsecciones='" & estcivil & "'")
            If Not rs.BOF Then
                codseccion = rs!resultado
            End If
            
            Set rs = cn.Execute("select codparticipantes from participante where usuario_cedusuario='" & cedula & "'")
            If Not rs.BOF Then
                cn.Execute ("insert into inscripcionucurricular VALUES(110," & codinscripcion & ",'" & PAIS1 & "'," & codparticipante & ",2," & codseccion & ",0,'')")
                Set rs = cn.Execute("select cantinscritos as resultado from inscripcion_seccion where secciones_codsecciones=" & codseccion)
                If Not rs.BOF Then
                    cant_insc = rs!resultado
                End If
                cn.Execute ("UPDATE INSCRIPCION_SECCION SET CANTINSCRITOS=" & cant_insc + 1 & " where secciones_codsecciones=" & codseccion)
            Else
                AppExcel.Application.Range("A" & fila).Select
                With AppExcel.Application.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    
                   
                End With
                cont = cont + 1
            End If
            fila = fila + 1
        Loop
        MsgBox "Listo Errores-->" & cont
    'End If
        
End Sub

Private Sub Command11_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    Dim codseccion As Integer
    Dim codparticipante As Long
    Dim codinscripcion As Long
    Dim cant_insc As Integer
    Dim trayecto As Integer
    Dim codpnf As Integer
    
    cont = 0
    
        'Set AppExcel = CreateObject("Excel.application")
        'AppExcel.application.Workbooks.Open FileName:=funciones.DirectorioActual & "migracion de inscritos 1-2014.xlsx"  'para abrir el libro
        'AppExcel.application.Windows("migracion de inscritos 1-2014.xlsx").Activate
        'AppExcel.application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
    codpnf = InputBox("Indique codigo del P.N.F:")
    trayecto = InputBox("Indique el trayecto:")
    seccion = InputBox("Indique la seccion:")
    cedula = InputBox("Indique la cedula:")
        'fila = 2
        'Do While AppExcel.application.cells(fila, 1).Value <> ""
            'cedula = AppExcel.application.cells(fila, 1).Value  'cedula
            'PAIS1 = AppExcel.application.cells(fila, 2).Value   'unidad curricular
            'estcivil = AppExcel.application.cells(fila, 3).Value    'seccion
            'codigo de participante
            Set rs = cn.Execute("select codparticipantes as resultado from participante where usuario_cedusuario='" & cedula & "'")
            If Not rs.BOF Then
                codparticipante = rs!resultado
            End If
            
            Set rsmat = cn.Execute("select codunidadcurricular as resultado from unidadcurricular where traunidadcurricular=" & trayecto & " and pnf_codpnf='" & codpnf & "'")
            'selecciona todas las unidades curriculares del trayecto
            If Not rsmat.BOF Then
                Do While Not rsmat.EOF
                    'codigo de inscripcion
                     Set rs = cn.Execute("select max(codinscripcionucurricular) as resultado from inscripcionucurricular")
                    If Not rs.BOF Then
                        codinscripcion = rs!resultado + 1
                    End If
                    'codigo de seccion
                    Set rs = cn.Execute("select codsecciones as resultado from secciones where unidadcurricular_codunidadcurricular='" & rsmat!resultado & "' and nomsecciones='" & seccion & "'")
                    If Not rs.BOF Then
                        codseccion = rs!resultado
                    End If
                
                    Set rs = cn.Execute("select codparticipantes from participante where usuario_cedusuario='" & cedula & "'")
                    If Not rs.BOF Then
                        cn.Execute ("insert into inscripcionucurricular VALUES(110," & codinscripcion & ",'" & rsmat!resultado & "'," & codparticipante & ",2," & codseccion & ",0,'')")
                        Set rs = cn.Execute("select cantinscritos as resultado from inscripcion_seccion where secciones_codsecciones=" & codseccion)
                        If Not rs.BOF Then
                            cant_insc = rs!resultado
                        End If
                        cn.Execute ("UPDATE INSCRIPCION_SECCION SET CANTINSCRITOS=" & cant_insc + 1 & " where secciones_codsecciones=" & codseccion)
                    Else
                        AppExcel.Application.Range("A" & fila).Select
                        With AppExcel.Application.Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            
                           
                        End With
                        cont = cont + 1
                    End If
                    rsmat.MoveNext
                Loop
            End If
            fila = fila + 1
        'Loop
        MsgBox "Listo Errores-->" & cont
    'End If
End Sub

Private Sub Command12_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    Dim codseccion As Integer
    Dim codparticipante As Long
    Dim codinscripcion As Long
    Dim cant_insc As Integer
    
    cont = 0
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "migracion de inscritos 1-2014.xlsx"  'para abrir el libro
        AppExcel.Application.Windows("migracion de inscritos 1-2014.xlsx").Activate
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            cedula = AppExcel.Application.Cells(fila, 1).Value  'cedula
            PAIS1 = AppExcel.Application.Cells(fila, 2).Value   'unidad curricular
            estcivil = AppExcel.Application.Cells(fila, 3).Value    'seccion
            'codigo de participante
             Set rs = cn.Execute("select codparticipantes as resultado from participante where usuario_cedusuario='" & cedula & "'")
            If Not rs.BOF Then
                codparticipante = rs!resultado
            End If
            'codigo de inscripcion
            Set rs = cn.Execute("select codinscripcionucurricular as resultado from inscripcionucurricular where participantes_codparticipantes=" & codparticipante & " and unidadcurricular_codunidadcurricular='" & PAIS1 & "' and periodosacademicos_codperiodosacademicos=110")
            If Not rs.BOF Then
                codinscripcion = rs!resultado
            End If
            'codigo de seccion
            Set rs = cn.Execute("select codsecciones as resultado from secciones where unidadcurricular_codunidadcurricular='" & PAIS1 & "' and nomsecciones='" & estcivil & "'")
            If Not rs.BOF Then
                codseccion = rs!resultado
            End If
            
            Set rs = cn.Execute("select codparticipantes from participante where usuario_cedusuario='" & cedula & "'")
            If Not rs.BOF Then
                cn.Execute ("update inscripcionucurricular set secciones_codsecciones=" & codseccion & " where codinscripcionucurricular=" & codinscripcion)
                Set rs = cn.Execute("select cantinscritos as resultado from inscripcion_seccion where secciones_codsecciones=" & codseccion)
                If Not rs.BOF Then
                    cant_insc = rs!resultado
                End If
                cn.Execute ("UPDATE INSCRIPCION_SECCION SET CANTINSCRITOS=" & cant_insc + 1 & " where secciones_codsecciones=" & codseccion)
            Else
                AppExcel.Application.Range("A" & fila).Select
                With AppExcel.Application.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    
                   
                End With
                cont = cont + 1
            End If
            fila = fila + 1
        Loop
        MsgBox "Listo Errores-->" & cont
    'End If
End Sub

Private Sub Command13_Click()
    Dim cont As Integer
    
    cont = 1
    Set rs = cn.Execute("select * from pensumunidadcurricular order by pensum_codpensum asc")
    If Not rs.BOF Then
        Do While Not rs.EOF
            cn.Execute ("update pensumunidadcurricular set codpensumunidadcurricular=" & cont & " where  unidadcurricular_codunidadcurricular='" & rs!unidadcurricular_codunidadcurricular & "' and pensum_codpensum=" & rs!pensum_codpensum)
            cont = cont + 1
            rs.MoveNext
        Loop
    End If
    MsgBox "Listo"
End Sub

Private Sub Command2_Click()
   ' On Error GoTo Error
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim fecingper As Date
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "periodosacademicos.xlsx"  'para abrir el libro
        AppExcel.Application.Windows("periodosacademicos.XLSX").Activate
        AppExcel.Application.Sheets("periodos").Select
        AppExcel.Application.Visible = True

        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            codper = AppExcel.Application.Cells(fila, 1).Value
            desper = AppExcel.Application.Cells(fila, 2).Value
            activo = AppExcel.Application.Cells(fila, 3).Value
            If activo = 1 Then
                cn.Execute ("insert into periodosacademicos values (" & codper & ",'" & desper & "','1900-01-01','1900-01-01',true)")
            Else
                cn.Execute ("insert into periodosacademicos values (" & codper & ",'" & desper & "','1900-01-01','1900-01-01',false)")
            End If
            barra = (fila / 110) * 100
            fila = fila + 1
        Loop
        MsgBox "Listo"

    
    
End Sub

Private Sub Command3_Click()
     Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim fecingper As Date
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "sno_familiar.xls"  'para abrir el libro
        AppExcel.Application.Windows("sno_familiar.XLS").Activate
        AppExcel.Application.Sheets("hoja1").Select
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal")
    'If Not rs.BOF Then
        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            codemp = CStr(AppExcel.Application.Cells(fila, 1).Value)
            codper = CStr(AppExcel.Application.Cells(fila, 2).Value)
            cedfam = CStr(AppExcel.Application.Cells(fila, 3).Value)
            nomfam = CStr(AppExcel.Application.Cells(fila, 4).Value)
            apefam = CStr(AppExcel.Application.Cells(fila, 5).Value)
            sexfam = CStr(AppExcel.Application.Cells(fila, 6).Value)
            fecnacfam = CDate(AppExcel.Application.Cells(fila, 7).Value)
            nexfam = CStr(AppExcel.Application.Cells(fila, 8).Value)
            estfam = CStr(AppExcel.Application.Cells(fila, 9).Value)
            hcfam = CStr(AppExcel.Application.Cells(fila, 10).Value)
            
            hcmfam = CStr(AppExcel.Application.Cells(fila, 11).Value)
            hijespfam = "0" 'CStr(AppExcel.application.cells(fila, 12).Value)
            
            estbonjug = CStr(AppExcel.Application.Cells(fila, 13).Value)
            cedula = "0" 'CStr(AppExcel.application.cells(fila, 14).Value)
            
            cn.Execute ("insert into sno_familiar values('" & codemp & "','" & codper & "','" & cedfam & "','" & nomfam & "','" & apefam & _
                "','" & sexfam & "','" & fecnacfam & "','" & nexfam & "','" & estfam & "','" & hcfam & "','" & hcmfam & _
                "','" & hijespfam & "','" & estbonjug & "','" & cedula & "')")
                
            barra = (fila / 1546) * 100
            fila = fila + 1
        Loop
        MsgBox "Listo"
    'End If
End Sub

Private Sub Command4_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "unidadcurricular.xlsx"  'para abrir el libro
        AppExcel.Application.Windows("unidadcurricular.xlsx").Activate
        AppExcel.Application.Visible = True
        
        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            codunidad = AppExcel.Application.Cells(fila, 1).Value
            desunidad = AppExcel.Application.Cells(fila, 2).Value
            trayecto = AppExcel.Application.Cells(fila, 3).Value
            
            pnf = AppExcel.Application.Cells(fila, 4).Value
            uc = AppExcel.Application.Cells(fila, 5).Value
            horas = AppExcel.Application.Cells(fila, 6).Value
            electiva = AppExcel.Application.Cells(fila, 7).Value
            mencion = AppExcel.Application.Cells(fila, 8).Value
            cn.Execute ("insert into unidadcurricular values('" & codunidad & "','" & desunidad & "'," & trayecto & "," & pnf & "," & uc & "," & horas & ",'" & electiva & "'," & mencion & ")")
            fila = fila + 1
        Loop
        MsgBox "Listo"
End Sub

Private Sub Command5_Click()
    ' On Error GoTo Error
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim fecingper As Date
    Dim codper As String
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "becas\CHEQUE.xls"  'para abrir el libro
        AppExcel.Application.Windows("CHEQUE.XLS").Activate
        AppExcel.Application.Sheets("hoja2").Select
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personalnomina")
    'If Not rs.BOF Then
        fila = 2
        Do While AppExcel.Application.Cells(fila, 2).Value <> ""
            separador = 0
            separador2 = 0
            cedula = AppExcel.Application.Cells(fila, 1).Value
            'cedula = Mid(AppExcel.application.cells(fila, 1).Value, 1, InStr(AppExcel.application.cells(fila, 1).Value, ".") - 1)
            nombrecomp = AppExcel.Application.Cells(fila, 2)
            separador = InStr(1, AppExcel.Application.Cells(fila, 2), " ")
            separador2 = InStr(separador + 1, AppExcel.Application.Cells(fila, 2), " ")
            If separador2 > 0 Then
                apellido = Mid(nombrecomp, 1, separador2 - 1)
                nombre = Mid(nombrecomp, separador2 + 1, Len(nombrecomp))
            Else
                apellido = Mid(nombrecomp, 1, separador - 1)
                nombre = Mid(nombrecomp, separador + 1, Len(nombrecomp))
            End If
            Set rs = cn.Execute("select * from rpc_beneficiario where ced_bene='" & cedula & "'")
            If rs.BOF Then
                cn.Execute ("insert into rpc_beneficiario (codemp,ced_bene,rifben,nombene, apebene, dirbene,sc_cuenta, nacben, codpai,codest,codmun,codpar,codbansig, tipconben,codtipcta,codban) values ('0001','" & cedula & "','V-" & cedula & "-0','" & nombre & "','" & apellido & "','VALENCIA','2110499005','V','058','---','---','---','','O','s1','---')")
                AppExcel.Application.Cells(fila, 7).Value = 1
            End If
            'MsgBox Cn.Errors
            
            barra = ((fila - 2) / 612) * 100
            fila = fila + 1
        Loop
        
        MsgBox "Listo"
    'End If
End Sub

Private Sub Command6_Click()
    Load Form2
    Form2.Show
End Sub

Private Sub Command7_Click()
       ' On Error GoTo Error
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim fecingper As Date
    Dim per As String
    

        'Set AppExcel = CreateObject("Excel.application")
        'AppExcel.application.Workbooks.Open FileName:=funciones.DirectorioActual & "personalisr.xls"  'para abrir el libro
        'AppExcel.application.Windows("personalisr.XLS").Activate
        'AppExcel.application.sheets("Hoja1").Select
        'AppExcel.application.Visible = True
        funciones.Conectar2
        
        'fila = 2
        'Do While AppExcel.application.cells(fila, 1).Value <> ""
            codemp = "0001" 'AppExcel.application.cells(fila, 1).Value
            codper = InputBox("coddigo personal:") 'AppExcel.application.cells(fila, 2).Value
            Set rs = cn.Execute("select codper as resultado from sno_personalisr where codper='" & codper & "'")
            If rs.BOF Then
                For i = 1 To 12
                    If i < 10 Then
                        per = "0" & CStr(i)
                    Else
                        per = CStr(i)
                    End If
                    cn.Execute ("insert into sno_personalisr ( codemp,codper , codisr, porisr, codconret) values('0001','" & codper & "','" & per & "',0,'')")
                Next i
            Else
                MsgBox "personal ya esta registrado"
            End If
            'barra = (fila / 886) * 100
            'fila = fila + 1
        'Loop
        MsgBox "Listo"
    'End If
    
End Sub

Private Sub Command8_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    
    cont = 0
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "migracion 6to semestre.xls"  'para abrir el libro
        AppExcel.Application.Windows("migracion 6to semestre.xls").Activate
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
        fila = 2
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            cedula = AppExcel.Application.Cells(fila, 1).Value
            PAIS1 = AppExcel.Application.Cells(fila, 2).Value
            estcivil = AppExcel.Application.Cells(fila, 3).Value
            nombre = AppExcel.Application.Cells(fila, 4).Value
            apellido = AppExcel.Application.Cells(fila, 5).Value
            'If apellido = "N" Then
            '    apellido = "3"
            'Else
            '    apellido = "1"
            'End If
            correo = AppExcel.Application.Cells(fila, 6).Value
            telefono = AppExcel.Application.Cells(fila, 7).Value
               
            Set rs = cn.Execute("select codparticipantes from participante where usuario_cedusuario='" & cedula & "'")
            
            If rs.BOF Then
                cn.Execute ("insert into participante VALUES('" & cedula & "'," & PAIS1 & "," & estcivil & "," & nombre & "," & apellido & "," & _
                correo & "," & telefono & ")")
            Else
                AppExcel.Application.Range("A" & fila).Select
                With AppExcel.Application.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    
                   
                End With
                cont = cont + 1
            End If
            fila = fila + 1
        Loop
        MsgBox "Listo Errores-->" & cont
    'End If
        
End Sub

'Private Sub Command2_Click()
'   ' On Error GoTo Error
'    Dim AppExcel As Object
'    Dim rs As ADODB.Recordset
'    Dim fila As Integer
'    Dim fecingper As Date
'
'        Set AppExcel = CreateObject("Excel.application")
'        AppExcel.application.Workbooks.Open FileName:=funciones.DirectorioActual & "periodosacademicos.xlsx"  'para abrir el libro
'        AppExcel.application.Windows("periodosacademicos.XLSX").Activate
'        AppExcel.application.sheets("periodos").Select
'        AppExcel.application.Visible = True'
'
'        fila = 2
'        Do While AppExcel.application.cells(fila, 1).Value <> ""
'            codper = AppExcel.application.cells(fila, 1).Value
'            desper = AppExcel.application.cells(fila, 2).Value
'            activo = AppExcel.application.cells(fila, 3).Value
'            If activo = 1 Then
'                Cn.Execute ("insert into periodosacademicos values (" & codper & ",'" & desper & "','1900-01-01','1900-01-01',true)")
'            Else
'                Cn.Execute ("insert into periodosacademicos values (" & codper & ",'" & desper & "','1900-01-01','1900-01-01',false)")
'            End If
'            barra = (fila / 110) * 100
'            fila = fila + 1
'        Loop
'        MsgBox "Listo"
'End Sub

'Private Sub Command3_Click()
'     Dim AppExcel As Object
'    Dim rs As ADODB.Recordset
'    Dim fila As Integer
'    Dim fecingper As Date
'
'        Set AppExcel = CreateObject("Excel.application")
'        AppExcel.application.Workbooks.Open FileName:=funciones.DirectorioActual & "sno_familiar.xls"  'para abrir el libro
'        AppExcel.application.Windows("sno_familiar.XLS").Activate
'        AppExcel.application.sheets("hoja1").Select
'        AppExcel.application.Visible = True
'
'    'Set rs = Cn.Execute("select * from sno_personal")
'    'If Not rs.BOF Then
'        fila = 2
'        Do While AppExcel.application.cells(fila, 1).Value <> ""
'            codemp = CStr(AppExcel.application.cells(fila, 1).Value)
'            codper = CStr(AppExcel.application.cells(fila, 2).Value)
'            cedfam = CStr(AppExcel.application.cells(fila, 3).Value)
'            nomfam = CStr(AppExcel.application.cells(fila, 4).Value)
'            apefam = CStr(AppExcel.application.cells(fila, 5).Value)
'            sexfam = CStr(AppExcel.application.cells(fila, 6).Value)
'            fecnacfam = CDate(AppExcel.application.cells(fila, 7).Value)
'            nexfam = CStr(AppExcel.application.cells(fila, 8).Value)
'            estfam = CStr(AppExcel.application.cells(fila, 9).Value)
'            hcfam = CStr(AppExcel.application.cells(fila, 10).Value)
'
'            hcmfam = CStr(AppExcel.application.cells(fila, 11).Value)
'            hijespfam = "0" 'CStr(AppExcel.application.cells(fila, 12).Value)
'
'            estbonjug = CStr(AppExcel.application.cells(fila, 13).Value)
'            cedula = "0" 'CStr(AppExcel.application.cells(fila, 14).Value)
'
'            Cn.Execute ("insert into sno_familiar values('" & codemp & "','" & codper & "','" & cedfam & "','" & nomfam & "','" & apefam & _
'                "','" & sexfam & "','" & fecnacfam & "','" & nexfam & "','" & estfam & "','" & hcfam & "','" & hcmfam & _
'                "','" & hijespfam & "','" & estbonjug & "','" & cedula & "')")
'
'            barra = (fila / 1546) * 100
'            fila = fila + 1
'        Loop
'        MsgBox "Listo"
'    'End If
'End Sub

Private Sub Command9_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim cont As Integer
    
    cont = 0
    
        Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=funciones.DirectorioActual & "notaslicquim108.xlsx"  'para abrir el libro
        AppExcel.Application.Windows("notaslicquim108.xlsx").Activate
        AppExcel.Application.Visible = True
        
    'Set rs = Cn.Execute("select * from sno_personal order by codper asc")
    'If Not rs.BOF Then
        fila = 1
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            codper = AppExcel.Application.Cells(fila, 1).Value
            codparticipante = AppExcel.Application.Cells(fila, 2).Value
            codunidadcurr = AppExcel.Application.Cells(fila, 4).Value
            nota = AppExcel.Application.Cells(fila, 5).Value
            condicion = AppExcel.Application.Cells(fila, 8).Value
            
                     
            Set rs = cn.Execute("select * from inscripcionucurricular where participantes_codparticipantes='" & codparticipante & "'")
            If Not rs.BOF Then
                cn.Execute ("update inscripcionucurricular set calificacion='" & nota & "', condicion='" & condicion & "' where participantes_codparticipantes=" & codparticipante & " and unidadcurricular_codunidadcurricular='" & codunidadcurr & "' and periodosacademicos_codperiodosacademicos=" & codper)
            Else
                AppExcel.Application.Range("A" & fila).Select
                With AppExcel.Application.Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    
                   
                End With
                cont = cont + 1
            End If
            fila = fila + 1
        Loop
        MsgBox "Listo Errores-->" & cont
    'End If
End Sub

Private Sub Form_Load()
    funciones.Conectar
    
End Sub
