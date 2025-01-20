VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMigrar 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   3120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMigrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim AppExcel As Object
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim libro As String
    Dim sql As String
    Dim periodo As String
    Dim coduc As String
    Dim nota As Integer
    Dim seccion As String
    Dim cedula As String
    
    Dialog1.ShowOpen
    Set AppExcel = CreateObject("Excel.application")
        AppExcel.Application.Workbooks.Open FileName:=Dialog1.FileName  'para abrir el libro
        
        Dialog1.InitDir = "C:\Users\USUARIO\Desktop\ESCRITORIO\vbpostgres 3.0\Formatos"
        
        libro = Mid(Dialog1.FileName, InStr(1, Dialog1.FileName, "Formatos\") + 9, Len(Dialog1.FileName))
        AppExcel.Application.Windows(libro).Activate
        'AppExcel.Application.Visible = True
        
        cedula = AppExcel.Application.Cells(1, 2).Value
        fila = 4
        Do While AppExcel.Application.Cells(fila, 1).Value <> ""
            coduc = AppExcel.Application.Cells(fila, 1).Value
            periodo = AppExcel.Application.Cells(fila, 5).Value
            nota = AppExcel.Application.Cells(fila, 3).Value
            seccion = AppExcel.Application.Cells(fila, 4).Value
            
            sql = "select codperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos='" & periodo & "'"
            codperiodo = funciones.CampoEntero(sql, cn)
            If codperiodo = 0 Then
                MsgBox "Error, no existe el periodo académico", vbCritical
                GoTo salir
            Else
                'consulta si existe la seccion
                If seccion = "" Then
                    MsgBox "Error, Sección nula", vbCritical
                    GoTo salir
                Else
                    sql = "select codperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos='" & periodo & "'"
                    codperiodo = funciones.CampoEntero(sql, cn)
                    
                    sql = "select codsecciones as resultado from secciones where unidadcurricular_codunidadcurricular='" & coduc & "' and periodosacademicos_codperiodoacademico=" & codperiodo & " and nomsecciones='" & seccion & "'"
                    codseccion = funciones.CampoEnteroLargo(sql, cn)
                    If codseccion = 0 Then  'en caso de no existir la seccion, procede a crearla
                        sql = "select mencion_codmencion as resultado from mencionunidadcurricular where unidadcurricular_codunidadcurricular='" & coduc & "'"
                        codmencion = funciones.CampoEntero(sql, cn)
                        sql = "select pensum_codpensum as resultado from mencion where codmencion=" & codmencion
                        codpensum = funciones.CampoEntero(sql, cn)
                        sql = "select max(codsecciones) as resultado from secciones"
                        codseccion = funciones.proximocodigoregistro(sql, cn)
                        
                        sql = "insert into secciones values (" & codseccion & ",1,'" & seccion & "',30," & codperiodo & ",'" & coduc & "','FALSE'," & codpensum & "," & codmencion & ")"
                        cn.Execute (sql)
                    End If
                    'determina el codigo del participante
                    sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & cedula & "'"
                    codparticipante = funciones.CampoString(sql, cn)
                    If codparticipante = "000" Then
                        MsgBox "Error, participante no registrado...", vbCritical
                        GoTo salir
                    Else
                        'valida que la calificacion no este registrada previamente
                        sql = "select codinscripcionucurricular as resultado from inscripcionucurricular where participantes_codparticipantes=" & codparticipante & " and periodosacademicos_codperiodosacademicos=" & periodo & " and unidadcurricular_codunidadcurricular='" & coduc & "' and secciones_codsecciones=" & codseccion
                        codinscripcion = funciones.CampoEnteroLargo(sql, cn)
                        If codinscripcion = 0 Then
                            sql = "select max(codinscripcionucurricular) as resultado from inscripcionucurricular"
                            codinscripcion = funciones.proximocodigoregistro(sql, cn)
                            sql = "insert into inscripcionucurricular values (" & codperiodo & "," & codinscripcion & ",'" & coduc & "'," & codparticipante & ",2," & codseccion & "," & nota & ",'Aprobado')"
                            cn.Execute (sql)
                        End If
                    End If
                End If
            End If
            fila = fila + 1
        Loop
        MsgBox "Proceso Concluido...", vbInformation
        'AppExcel.Application.Saved = True
        AppExcel.Application.Quit
        Set AppExcel = Nothing
salir:
    
End Sub
