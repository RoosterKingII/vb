Attribute VB_Name = "funciones"
Public cn As ADODB.Connection
Public cn2 As ADODB.Connection
Public BdTecnologico As ADODB.Connection
Public BdPolitecnico As ADODB.Connection
Public regpromocion As Integer 'determina si esta registrado en una promocion, o no, ademas si se desean ingresar los datos manualmente
Public trainicio As String
Public transicion As String
Public totaluccur As Integer
Public totalnotauc As Integer
Public cedusuario As String
Public codperfil As Integer
Public desperfil As String
Public usuario As String
Public AppExcel As Object
Public codparticipante As Long
Public codpensum As Integer
Public codcohorte As String

Function DirectorioActual()
    ChDrive App.Path
    ChDir App.Path
    DirectorioActual = App.Path
    If Len(DirectorioActual) > 3 Then
        DirectorioActual = DirectorioActual & "\"
    End If
End Function



Sub Conectar()


    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=MSDASQL;" & _
                          "Driver={PostgreSQL ANSI};" & _
                          "SERVER=10.100.100.22;" & _
                          "DATABASE=GES;" & _
                          "UID=postgres;" & _
                          "PWD=admin1iuteval;"

    'Cn.ConnectionString = "DRIVER=PostGreSQL; Server=172.31.0.6;Port=5432;UserId=potsgres;" & _
    '                      "Password=admin1iuteval;Database=troya;"
    cn.Open
    If cn.State = 0 Then
        MsgBox "Error en Conexión a la Base de Datos", vbCritical
    End If
End Sub

Sub Conectar2()


    Set cn2 = New ADODB.Connection
    'cn.ConnectionString = "Provider=MSDASQL;" & _
                          "Driver={PostgreSQL ANSI};" & _
                          "SERVER=192.168.10.9;" & _
                          "DATABASE=mysql;" & _
                          "UID=postgres;" & _
                          "PWD=admin1iuteval;"

    'Cn.ConnectionString = "DRIVER=PostGreSQL; Server=172.31.0.6;Port=5432;UserId=potsgres;" & _
    '                      "Password=admin1iuteval;Database=troya;"
    'cn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
    '                      "SERVER=192.168.10.9;" & _
    '                      "DATABASE=censo;" & _
    '                      "User=root;" & _
    '                      "Password=admin1iuteval;"
        Set cn2 = New ADODB.Connection
    cn2.ConnectionString = "Provider=MSDASQL;" & _
                          "Driver={PostgreSQL ANSI};" & _
                          "SERVER=10.100.100.22;" & _
                          "DATABASE=ACC;" & _
                          "UID=postgres;" & _
                          "PWD=admin1iuteval;"
'.168.10.9;"
    cn2.Open
    If cn2.State = 0 Then
        MsgBox "Error en Conexión a la Base de Datos", vbCritical
    End If
End Sub

Sub Conectar3()


    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=MSDASQL;" & _
                          "Driver={PostgreSQL ANSI};" & _
                          "SERVER=192.168.10.9;" & _
                          "DATABASE=prueba2014;" & _
                          "UID=postgres;" & _
                          "PWD=admin1iuteval;"

    cn.Open
    If cn.State = 0 Then
        MsgBox "Error en Conexión a la Base de Datos", vbCritical
    End If
End Sub

Function CodPeriodoAcademicopnf(periodo As String) As Integer
    Conectar
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT codperiodosacademicos From periodosacademicos WHERE desperiodosacademicos='" & periodo & "'")
    If Not rs.BOF Then
        CodPeriodoAcademicopnf = rs!codperiodosacademicos
    Else
        CodPeriodoAcademicopnf = 0
    End If
    cn.Close
End Function

Function CodPnfpnf(despnf As String) As Integer
    Conectar
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("SELECT codpnf From pnf WHERE despnf='" & despnf & "'")
    If Not rs.BOF Then
        CodPnfpnf = rs!codpnf
    Else
        Codnfpnf = 0
    End If
    cn.Close
End Function

'/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*'
'FUNCIONES SQL SERVER
Public Function AbrirBdPolitecnico() As Boolean
Set BdPolitecnico = New Connection
BdPolitecnico.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=politecnico;Data Source=localhost"
BdPolitecnico.CursorLocation = adUseClient
BdPolitecnico.Open
If BdPolitecnico.State = 0 Then
    MsgBox "Error en Conexión a la Base de Datos Politecnico. Consulte al Administrador,", vbCritical, "Error en Conexión"
    AbrirBdPolitecnico = False
Else
    AbrirBdPolitecnico = True
End If
End Function

Public Function AbrirBdTecnologico() As Boolean
Set BdTecnologico = New Connection
BdTecnologico.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=tecnologico;Data Source=10.130.130.47"
BdTecnologico.CursorLocation = adUseClient
BdTecnologico.Open
If BdTecnologico.State = 0 Then
    MsgBox "Error en Conexión a la Base de Datos Phoenix. Consulte al Administrador,", vbCritical, "Error en Conexión"
    AbrirBdTecnologico = False
Else
    AbrirBdTecnologico = True
End If
End Function

Function alumnostraysempnf(pnf As String, codper As Integer, traysem As Integer) As Integer
    AbrirBdPolitecnico
    Dim rs As ADODB.Recordset
    Set rs = BdPolitecnico.Execute("SELECT COUNT(DISTINCT Ced_Alum) AS cuenta From dbo.[" & pnf & "] WHERE (Cod_Per = " & codper & ") AND (RIGHT(LEFT(Cod_Materia, 2), 1) =" & traysem & ")")
    If Not rs.BOF Then
        alumnostraysempnf = rs!cuenta
    Else
       MsgBox "no hay materias/unidad curricular para el semestre/trayecto seleccionado"
    End If
End Function

Function CampoString(sql As String, cn As Connection) As String
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        If Not IsNull(rs!resultado) Then
            CampoString = rs!resultado
        Else
            CampoString = "000"
        End If
    Else
        CampoString = "000"
    End If
End Function

Function CampoDouble(sql, cn As Connection) As Double
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        If IsNull(rs!resultado) Then
            CampoDouble = 0
        Else
            CampoDouble = rs!resultado
        End If
    Else
        CampoDouble = 0
    End If
End Function

Function CampoEntero(sql As String, cn As Connection) As Integer
    Dim rs As ADODB.Recordset
    
    If cn.State = 0 Then
        funciones.Conectar
    End If
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        If IsNull(rs!resultado) Then
            CampoEntero = 0
        Else
            CampoEntero = rs!resultado
        End If
    Else
        CampoEntero = 0
    End If
End Function

Function CampoEnteroLargo(sql As String, cn As Connection) As Long
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        CampoEnteroLargo = rs!resultado
    Else
        CampoEnteroLargo = 0
    End If
End Function

Function CampoBooleano(sql As String, cn As Connection) As Boolean
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        CampoBooleano = rs!resultado
    Else
        CampoBooleano = "FALSE"
    End If
End Function

Function ExisteRegistro(sql As String, cn As Connection) As Boolean
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        ExisteRegistro = True
    Else
        ExisteRegistro = False
    End If
End Function

Sub migrarnotassemestre(ced As String, esp As String, cn2 As ADODB.Connection, historico As Boolean, fila As Integer, AppExcel As Object)
    Dim rs As ADODB.Recordset
    Dim sql As String
    funciones.Conectar
    
    sql = "select codparticipantes as resultado from participante, participantecohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                      " participantecohorte.actparticipantecohorte='true' and participante.usuario_cedusuario='" & ced & "'"
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
      codparticipante = funciones.CampoEntero(sql, cn)
      If historico Then
        sql = "select hist_notas.cod_materia as codmateria, hist_notas.sec_mat as seccion, periodos.per_insc as periodo, hist_notas.nota_mat as nota, hist_notas.condicion as condicion from hist_notas, periodos" & _
              " where hist_notas.cod_per=periodos.cod_per and  hist_notas.ced_alum='" & ced & "' order by hist_notas.cod_materia asc"
      Else
        sql = "select [" & esp & "].cod_materia as codmateria, [" & esp & "].sec_mat as seccion, periodos.per_insc as periodo, [" & esp & "].nota_mat as nota, [" & esp & "].condicion as condicion from [" & esp & "], periodos" & _
                           " where [" & esp & "].cod_per=periodos.cod_per and  [" & esp & "].ced_alum='" & ced & "' order by [" & esp & "].cod_materia asc"
      End If
      Set rs = cn2.Execute(sql)
      If Not rs.BOF Then
        Do While Not rs.EOF
            sql = "select codperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos='" & rs!periodo & "'"
            codper = funciones.CampoEntero(sql, cn)
            sql = "select codsecciones as resultado from secciones where nomsecciones='" & rs!seccion & "' and periodosacademicos_codperiodoacademico=" & codper & " and unidadcurricular_codunidadcurricular='" & rs!codmateria & "'"
            codsec = funciones.CampoEntero(sql, cn)
            sql = "select max(codinscripcionucurricular) as resultado from inscripcionucurricular"
            codinsc = funciones.CampoEnteroLargo(sql, cn) + 1
            If rs!condicion <> "Aprobado" And rs!condicion <> "Reprobado" Then
                If IsNumeric(rs!nota) Then
                    If rs!nota >= 10 Then
                        condicion = "Aprobado"
                    Else
                        condicion = "Reprobado"
                    End If
                    nota = rs!nota
                Else
                    If rs!nota = "A" Then
                        condicion = "Aprobado"
                        nota = 20
                    ElseIf rs!nota = "EQ" Then
                        condicion = "Aprobado"
                        nota = 10
                    Else
                        condicion = "Reprobado"
                        nota = 1
                    End If
                End If
            Else
                If IsNumeric(rs!nota) Then
                    If rs!nota >= 10 Then
                        condicion = "Aprobado"
                    Else
                        condicion = "Reprobado"
                    End If
                    nota = rs!nota
                Else
                    If rs!nota = "A" Then
                        condicion = "Aprobado"
                        nota = 20
                    ElseIf rs!nota = "EQ" Then
                        condicion = "Aprobado"
                        nota = 10
                    Else
                        condicion = "Reprobado"
                        nota = 1
                    End If
                End If
            End If
            If codsec = 0 Then
                sql = "select max(codsecciones) as resultado from secciones"
                codsec = funciones.CampoEntero(sql, cn) + 1
                turno = Mid(rs!seccion, 2, 1)
                If Not IsNumeric(turno) Then
                    turno = 3
                Else
                    If CInt(turno) > 3 Then: turno = 3
                End If
                sql = "select desunidadcurricular as resultado from unidadcurricular where codunidadcurricular='" & rs!codmateria & "'"
                If funciones.ExisteRegistro(sql, cn) Then
                    sql = "insert into secciones values(" & codsec & "," & turno & ",'" & rs!seccion & "',50," & codper & ",'" & rs!codmateria & "')"
                    cn.Execute (sql)
                End If
            End If
            sql = "select desunidadcurricular as resultado from unidadcurricular where codunidadcurricular='" & rs!codmateria & "'"
            If funciones.ExisteRegistro(sql, cn) Then
                sql = "select codinscripcionucurricular as resultado from inscripcionucurricular where participantes_codparticipantes=" & codparticipante & " and unidadcurricular_codunidadcurricular='" & rs!codmateria & "' and" & _
                    " periodosacademicos_codperiodosacademicos=" & codper & " and secciones_codsecciones=" & codsec
                If Not funciones.ExisteRegistro(sql, cn) Then
                    sql = "insert into inscripcionucurricular values(" & codper & "," & codinsc & ",'" & rs!codmateria & "'," & codparticipante & ",1," & codsec & "," & nota & ",'" & condicion & "')"
                    cn.Execute (sql)
                End If
            End If
            rs.MoveNext
        Loop
        AppExcel.Application.Range("A" & fila).Select
        With AppExcel.Application.Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = vbGreen
        End With
        AppExcel.Application.Cells(fila, 9).Value = 3
      'End If
        'MsgBox "Proceso concluido", vbInformation
      Else
        AppExcel.Application.Range("A" & fila).Select
        With AppExcel.Application.Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = vbBlue
        End With
        AppExcel.Application.Cells(fila, 9).Value = 2
      End If
    Else
        If MsgBox("no esta registrado como participante, desea migrarlo?", vbQuestion + vbYesNo) = vbYes Then
            sql = "select alumno.nom_alum as resultado from alumno, especialidad, periodos " & _
                  " where alumno.cod_esp=especialidad. cod_esp and alumno.cod_per=periodos.cod_per and alumno.ced_alum='" & ced & "'"
            nomusuario = funciones.CampoString(sql, cn2)
            sql = "select alumno.ape_alum as resultado from alumno, especialidad, periodos " & _
                  " where alumno.cod_esp=especialidad. cod_esp and alumno.cod_per=periodos.cod_per and alumno.ced_alum='" & ced & "'"
            apeusuario = funciones.CampoString(sql, cn2)
            sql = "select especialidad.des_esp as resultado from alumno, especialidad, periodos " & _
                  " where alumno.cod_esp=especialidad. cod_esp and alumno.cod_per=periodos.cod_per and alumno.ced_alum='" & ced & "'"
            especialidad = funciones.CampoString(sql, cn2)
            Select Case especialidad
                Case "Informática": especialidad = "INFORMATICA (SEMESTRAL)"
                Case "Polímeros": especialidad = "POLIMEROS"
                Case "Electricidad": especialidad = "ELECTRICIDAD (SEMESTRAL)"
                Case "Química": especialidad = "QUIMICA (SEMESTRAL)"
                Case "Química (A)": especialidad = "QUIMICA (ANUAL)"
                Case "Electricidad (A)": especialidad = "ELECTRICIDAD (ANUAL)"
                Case "Polímeros(V)": especialidad = "POLIMEROS (PENSUM VIEJO)"
            End Select
            sql = "select alumno.sexo as resultado from alumno, especialidad, periodos " & _
                  " where alumno.cod_esp=especialidad. cod_esp and alumno.cod_per=periodos.cod_per and alumno.ced_alum='" & ced & "'"
            sexo = funciones.CampoString(sql, cn2)
            If sexo = "F" Then
                sexo = "1"
            Else
                sexo = "0"
            End If
            
            sql = "select periodos.per_insc as resultado from alumno, especialidad, periodos " & _
                  " where alumno.cod_esp=especialidad. cod_esp and alumno.cod_per=periodos.cod_per and alumno.ced_alum='" & ced & "'"
            periodo = funciones.CampoString(sql, cn2)
            Conectar
            sql = "select nomapeusuario as resultado  from usuario where cedusuario='" & ced & "'"
            If Not funciones.ExisteRegistro(sql, cn) Then
                sql = "insert into usuario values('" & ced & "','058',1,'" & nomusuario & "','" & apeusuario & "','','','','1900-01-01','','300001'," & _
                    "'3001','3001','300001','30','30',0,'" & sexo & "','058','000','000','" & nomusaurio & " " & apeusuario & "','FALSE',0,'FALSE','N/A','C',0)"
                cn.Execute (sql)
            End If
            sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & ced & "'"
            If Not funciones.ExisteRegistro(sql, cn) Then
                sql = "select max(codparticipantes) as resultado from participante"
                codparticipante = funciones.CampoEnteroLargo(sql, cn) + 1
                sql = "insert into participante values('" & ced & "'," & codparticipante & ")"
                cn.Execute (sql)
            Else
                codparticipante = funciones.CampoEnteroLargo(sql, cn)
            End If
            sql = "select codparticipantecohorte as resultado from participante, participantecohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and participante.usuario_cedusuario='" & ced & "'"
            If Not funciones.ExisteRegistro(sql, cn) Then
                sql = "select cohorte.codcohorte as resultado from cohorte,pnf where cohorte.pnf_codpnf=pnf.codpnf and pnf.despnf='" & especialidad & "'"
                codcohorte = funciones.CampoEntero(sql, cn)
                'especialidad = "QUIMICA(ANUAL)"
                sql = "select codperiodosacademicos as resultado from periodosacademicos where desperiodosacademicos='" & periodo & "'"
                codperiodo = funciones.CampoEntero(sql, cn)
                sql = "select max(codparticipantecohorte) as resultado from participantecohorte"
                codparticipantecohorte = funciones.CampoEnteroLargo(sql, cn) + 1
                sql = "insert into participantecohorte values(" & codparticipantecohorte & "," & codparticipante & "," & codcohorte & ",1,1," & codperiodo & ",2,TRUE)"
                cn.Execute (sql)
            End If
        End If
    
    End If
    
End Sub

Sub llenarcombobox(objeto As ComboBox, sql As String, cn As ADODB.Connection, estatus As Boolean)
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If Not rs.BOF Then
        objeto.Clear
        objeto.AddItem ("Seleccione")
        Do While Not rs.EOF
            objeto.AddItem rs!resultado
            rs.MoveNext
        Loop
        If estatus Then
            objeto.Text = objeto.List(0)
        End If
    End If
End Sub

Function Nivel_Participante(codparticipante As Long) As Integer
    Dim rs1 As ADODB.Recordset
    Dim sql As String
    
    Dim codpensum As Integer
    Dim codcohorte As Integer
    Dim trainicio As Integer
    Dim transicion As Integer
    Dim cedula As String
  
    sql = "select usuario_cedusuario as resultado from participante where participante.codparticipantes='" & codparticipante & "'"
    cedula = funciones.CampoEnteroLargo(sql, cn)
    
    sql = "select pensum.codpensum as resultado  from participante,participantecohorte, cohorte, pnf,pensum where participante.codparticipantes = participantecohorte.participante_codparticipante AND" & _
          " participantecohorte.cohorte_codcohorte = cohorte.codcohorte AND cohorte.pnf_codpnf = pnf.codpnf AND cohorte.pensum_codpensum = pensum.codpensum AND participante.usuario_cedusuario = '" & cedula & "' and" & _
          " participantecohorte.actparticipantecohorte='TRUE'"
    codpensum = funciones.CampoEntero(sql, cn)
  
    sql = "select cohorte.codcohorte as resultado  from participante,participantecohorte, cohorte where participante.codparticipantes = participantecohorte.participante_codparticipante AND" & _
          " participantecohorte.cohorte_codcohorte = cohorte.codcohorte AND participante.usuario_cedusuario = '" & cedula & "' and cohorte.pensum_codpensum=" & codpensum & " and participantecohorte.actparticipantecohorte='TRUE'"
    codcohorte = funciones.CampoEntero(sql, cn)
    
    sql = "select trapensumnivel as resultado from cohorte, nivel, pensumnivel where cohorte.nivel_codnivel = nivel.codnivel AND nivel.codnivel = pensumnivel.nivel_codnivel AND" & _
          " pensumnivel.pensum_codpensum = " & codpensum & " AND cohorte.codcohorte =" & codcohorte
    trainicio = funciones.CampoEntero(sql, cn)
  
    'sql = "select codparticipantes as resultado from participante where usuario_cedusuario='" & cedula & "'"
    'codparticipante = funciones.CampoEnteroLargo(sql, Cn)
  
    If trainicio = 1 Then 'inicio desde T0
        sql = "select max(trapensumnivel) as resultado from cohorte, nivel, pensumnivel where cohorte.nivel_codnivel = nivel.codnivel AND nivel.codnivel = pensumnivel.nivel_codnivel AND" & _
              " pensumnivel.pensum_codpensum = " & codpensum
        transicion = funciones.CampoEntero(sql, cn)
        sql = "select count(inscripcionucurricular) as resultado from inscripcionucurricular, participante, unidadcurricular where participante.codparticipantes = inscripcionucurricular.participantes_codparticipantes AND" & _
              " inscripcionucurricular.unidadcurricular_codunidadcurricular = unidadcurricular.codunidadcurricular AND  participante.usuario_cedusuario = '" & cedula & "' and unidadcurricular.traunidadcurricular>='" & transicion & "' and calificacion>0"
        If funciones.CampoEntero(sql, cn) > 0 Then  ' si tiene notas despues de la transcicion es TSU, nivel=2
            condicion1 = " and cast(unidadcurricular.traunidadcurricular AS int) >=" & trainicio & " and unidadcurricular_codunidadcurricular not in(SELECT codunidadcurricular FROM unidadcurricular where traunidadcurricular = '0" & transicion & "')"
            sql = "select trapensum as resultado from pensum where codpensum=" & codpensum
            transicion = funciones.CampoEntero(sql, cn)
            nivel = 2
        Else        'es bachiller nivel=1
            nivel = 1
        End If
    Else                   'inicio en transicion o posterior"
        sql = "select trapensum as resultado from pensum where codpensum=" & codpensum
        transicion = funciones.CampoEntero(sql, cn)
        sql = "select count(participantecohorte.codparticipantecohorte) as resultado FROM participantecohorte, participante, cohorte where participante.codparticipantes = participantecohorte.participante_codparticipante AND participantecohorte.cohorte_codcohorte = cohorte.codcohorte AND" & _
              " participante.usuario_cedusuario = '" & cedula & "' and substring(cohorte.descohorte from 1 for 2)<>'PS'"
    
        If funciones.CampoEntero(sql, cn) = 0 Then  '
            nivel = 2
        Else                                        ' este es el caso de cambio de pnf (mateiales a materiales ind) y los cambios de malla en Pq y elect
            sql = "select cohorte.pensum_codpensum as resultado from participante, participantecohorte, cohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participantecohorte.actparticipantecohorte='FALSE' and participante.usuario_cedusuario='" & cedula & "'"
            codpnffalso = funciones.CampoEntero(sql, cn)
            
            sql = "select cohorte.pensum_codpensum as resultado from participante, participantecohorte, cohorte where participante.codparticipantes=participantecohorte.participante_codparticipante and" & _
                " participantecohorte.cohorte_codcohorte=cohorte.codcohorte and participantecohorte.actparticipantecohorte='TRUE' and participante.usuario_cedusuario='" & cedula & "'"
            codpnfverdad = funciones.CampoEntero(sql, cn)
            If codpnffalso = codpnfverdad And Len(CStr(codpnffalso)) > 0 Then
                nivel = 2
            Else
                nivel = 1
            End If
        End If
    End If
    
    Nivel_Participante = nivel
End Function

Function CalcularIAG(cedula As String, codpensum As Integer) As Double

  Dim NOTAXUC As Integer
  Dim sumUC As Integer
  Dim i As Integer
  Dim indice As Double
  Dim sql As String
  Dim sql1 As String
  Dim rs As ADODB.Recordset
  
  sql = "select unidadcurricular_codunidadcurricular as resultado from pensumunidadcurricular where pensum_codpensum=" & codpensum
  Set rs = cn.Execute(sql)
  If Not rs.BOF Then
    NOTAXUC = 0
    sumUC = 0
    Do While Not rs.EOF
        sql = "select max(calificacion) as resultado from inscripcionucurricular, participante where participantes_codparticipantes=participante.codparticipantes and" & _
              " inscripcionucurricular.unidadcurricular_codunidadcurricular='" & rs!resultado & "' and participante.usuario_cedusuario='" & cedula & "' and substring(inscripcionucurricular.condicion from 1 for 4)='Apro'"
        sql1 = "select ucunidadcurricular as resultado from unidadcurricular where codunidadcurricular='" & rs!resultado & "'"
        If Not rs.BOF Then
            NOTAXUC = NOTAXUC + funciones.CampoEntero(sql, cn) * funciones.CampoEntero(sql1, cn)
            If funciones.CampoEntero(sql, cn) > 0 Then  'si nota es cero entonces no curso y no se suman las uc
                sumUC = sumUC + funciones.CampoEntero(sql1, cn)
            End If
        End If
        rs.MoveNext
    Loop
  End If
    If sumUC = 0 Then
        indice = 0
    Else
        indice = NOTAXUC / sumUC
    End If
  CalcularIAG = indice
End Function

Function proximocodigoregistro(sql As String, cn As ADODB.Connection) As Long
    Dim rs As ADODB.Recordset
    
    Set rs = cn.Execute(sql)
    If rs.BOF Then
        proximocodigoregistro = 1
    Else
        If IsNull(rs!resultado) Then
            proximocodigoregistro = 1
        Else
            proximocodigoregistro = rs!resultado + 1
        End If
    End If
End Function

Function fechacompleta(fechacorta As String) As String
    Dim dia As String
    Dim mes As String
    Dim anno As String
    Dim fechaaux As String
    
    dia = Mid(fechacorta, 1, 2)
    mes = Mid(fechacorta, 4, 2)
    anno = Mid(fechacorta, 7, 4)
    
    fechaaux = fechaaux & dia
    Select Case mes
        Case "01": fechaaux = fechaaux & " de " & "Enero"
        Case "02": fechaaux = fechaaux & " de " & "Febrero"
        Case "03": fechaaux = fechaaux & " de " & "Marzo"
        Case "04": fechaaux = fechaaux & " de " & "Abril"
        Case "05": fechaaux = fechaaux & " de " & "Mayo"
        Case "06": fechaaux = fechaaux & " de " & "Junio"
        Case "07": fechaaux = fechaaux & " de " & "Julio"
        Case "08": fechaaux = fechaaux & " de " & "Agosto"
        Case "09": fechaaux = fechaaux & " de " & "Septiembre"
        Case "10": fechaaux = fechaaux & " de " & "Octubre"
        Case "11": fechaaux = fechaaux & " de " & "Noviembre"
        Case "12": fechaaux = fechaaux & " de " & "Diciembre"
    End Select
    fechaaux = fechaaux & " de " & anno
    fechacompleta = fechaaux
End Function

Function formatofolio(numero As Integer) As String
    If numero > 0 And numero < 10 Then
        formatofolio = "00" & numero
    ElseIf numero >= 10 And numero < 100 Then
        formatofolio = "0" & numero
    Else
        formatofolio = numero
    End If
End Function

Function formatocorrelativo(numero As Integer) As String
    If numero > 0 And numero < 10 Then
        formatocorrelativo = "000" & numero
    ElseIf numero >= 10 And numero < 100 Then
        formatocorrelativo = "00" & numero
    ElseIf numero >= 100 And numero <= 1000 Then
        formatocorrelativo = "0" & numero
    Else
        formatocorrelativo = numero
    End If
End Function

Sub ENCABEZADOSEMESTRE(fil As Integer, semestre As String, regimen As String)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A7:I8").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
    If Len(semestre) > 1 Then
        If CInt(Mid(semestre, 2, 1)) > 1 Then
            semestre = "Transición"
        Else
            semestre = "Inicial"
        End If
    End If
    semestre = regimen & " " & semestre
    AppExcel.Application.Cells(fil, 1).Value = semestre
End Sub

Sub INSERTARFILA(fil As Integer)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A9:I9").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
End Sub

Sub PIESEMESTRE(fil As Integer)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A10:H11").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
End Sub

Sub PIEDOCUMENTO(fil As Integer)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A13:J20").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
    
    'AppExcel.Application.Sheets(1).Select
    'AppExcel.Application.Range("A14:J15").Select
    'AppExcel.Application.Selection.Copy
    'AppExcel.Application.Sheets("CERTIFICADAS1").Select
    'AppExcel.Application.Range("A" & fil + 2).Select
    'AppExcel.Application.ActiveSheet.Paste
    
    'AppExcel.Application.Sheets(1).Select
    'AppExcel.Application.Range("A16:J17").Select
    'AppExcel.Application.Selection.Copy
    'AppExcel.Application.Sheets("CERTIFICADAS1").Select
    'AppExcel.Application.Range("A" & fil + 4).Select
    'AppExcel.Application.ActiveSheet.Paste
    
    'AppExcel.Application.Sheets(1).Select
    'AppExcel.Application.Range("A18:J19").Select
    'AppExcel.Application.Selection.Copy
    'AppExcel.Application.Sheets("CERTIFICADAS1").Select
    'AppExcel.Application.Range("A" & fil + 7).Select
    'AppExcel.Application.ActiveSheet.Paste
End Sub

Sub COLOCARNOTA(trayecto As String, codparticipante As Long, ByRef pag As Integer, ByRef pagina As Integer, ByRef fil As Integer, control As String)
    Dim rs As ADODB.Recordset
    Set rs = cn.Execute("select unidadcurricular.desunidadcurricular as materia, unidadcurricular.ucunidadcurricular as uc, inscripcionucurricular.calificacion as nota, periodosacademicos.desperiodosacademicos as periodo" & _
                        " From inscripcionucurricular, unidadcurricular, periodosacademicos, pnf, pensumunidadcurricular where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and inscripcionucurricular.periodosacademicos_codperiodosacademicos=periodosacademicos.codperiodosacademicos and" & _
                        " inscripcionucurricular.unidadcurricular_codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and unidadcurricular.pnf_codpnf=pnf.codpnf and pensumunidadcurricular.pensum_codpensum=" & codpensum & " and pnf.despnf='" & FrmConstancias.txtespecialidad.Text & "' and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & _
                        " and unidadcurricular.traunidadcurricular='" & trayecto & "' and substring(inscripcionucurricular.condicion from 1 for 4)='Apro' order by materia asc")
    If Not rs.BOF Then
       Do While Not rs.EOF
        If fil > 59 Then
            'vectorimpresion(indice) = "CERTIFICADAS" & pag
            'indice = indice + 1
            pagina = pagina + 1
            pag = pag + 59
            fil = 1
            funciones.Encabezado_Hoja FrmConstancias.txtcedula.Text, FrmConstancias.txtapellidos.Text & ", " & FrmConstancias.txtnombres.Text, FrmConstancias.txtespecialidad.Text, control, pagina, fil, pag
            fil = fil + 6
        End If
          funciones.INSERTARFILA pag + fil - 1
          
          AppExcel.Application.Cells(pag + fil - 1, 1).Value = rs!materia
          If rs!nota = "A" Then
              AppExcel.Application.Cells(pag + fil - 1, 5).Value = rs!nota
              AppExcel.Application.Cells(pag + fil - 1, 6).Value = "Aprobado"
          Else
              AppExcel.Application.Cells(pag + fil - 1, 5).Value = rs!nota
              AppExcel.Application.Cells(pag + fil - 1, 6).Value = funciones.EnLetras(rs!nota)
          End If
          AppExcel.Application.Cells(pag + fil - 1, 7).Value = rs!uc
          AppExcel.Application.Cells(pag + fil - 1, 8).Value = funciones.EnLetras(rs!uc)
          AppExcel.Application.Cells(pag + fil - 1, 9).Value = rs!periodo
          fil = fil + 1
          'If funciones.SALTODEPAGINA(fil) = True Then
          '   AppExcel.Application.activesheet.PageSetup.printarea = "$A$1:$H$" & fil
          '   fil = 7
          '   pag = 2
          '   AppExcel.Application.sheets(pag + 1).Select
          'End If
          rs.MoveNext
       Loop
    End If
End Sub

Function RESUMENPERIODO(trayecto As String, codparticipante As Long, pag As Integer, fila As Integer)
    
    With AppExcel.Application
        .Sheets("CERTIFICADAS1").Cells(fila, 6) = funciones.UCCURSADAS(trayecto, codparticipante) & "  (" & funciones.EnLetras(funciones.UCCURSADAS(trayecto, codparticipante)) & ")"
        fila = fila + 1
        notauc = funciones.NOTAXUC(trayecto, codparticipante)
        uccur = funciones.UCCURSADAS(trayecto, codparticipante)
        totalnotauc = totalnotauc + notauc
        totaluccur = totaluccur + uccur
        If uccur <> 0 Then
            iraa = Round(notauc / uccur, 2)
        Else
            iraa = 0
        End If
        AppExcel.Application.Cells(fila, 6).Value = ConvertirNumerosEnLetras(CStr(iraa))
    End With
End Function

Sub RESUMENTOTAL(codparticipante As Long, fila As Integer)
    'coloca el total de unidades credito cursadas y aprobadas, el IAG y el IAG en la hoja de resumen en caso de no estar registrado en una promocion
    With AppExcel.Application
        totalnotauc = totalnotauc + notauc
        totaluccur = totaluccur + uccur 'acumula los creditos tanto de tsu, como de ingeniero o licenciado
        .Sheets("CERTIFICADAS1").Cells(fila, 7) = totaluccur & "  (" & funciones.EnLetras(CStr(totaluccur)) & ")"
        fila = fila + 1
        If totaluccur <> 0 Then
            iraa = Round(totalnotauc / totaluccur, 2)
        Else
            iraa = 0
        End If
        AppExcel.Application.Cells(fila, 7).Value = ConvertirNumerosEnLetras(CStr(iraa))
    End With
End Sub

Function UCCURSADAS(trayecto As String, codparticipante As Long) As Integer
    Dim rs As ADODB.Recordset
    'Set rs = cn.Execute("select sum(unidadcurricular.ucunidadcurricular) as resultado from pensumunidadcurricular, unidadcurricular where pensumunidadcurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and" & _
    '                    " pensum_codpensum=" & codpensum & " and unidadcurricular.traunidadcurricular='" & trayecto & "'")
    '
    Set rs = cn.Execute("select sum(unidadcurricular.ucunidadcurricular) as resultado from inscripcionucurricular, unidadcurricular, pnf, pensumunidadcurricular" & _
            " where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
            " pensumunidadcurricular.pensum_codpensum=" & codpensum & " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & FrmConstancias.txtespecialidad.Text & "' and unidadcurricular.traunidadcurricular='" & trayecto & "'" & _
            " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and substring(inscripcionucurricular.condicion from 1 for 4)='Apro'")
    If Not rs.BOF Then
        If Not IsNull(rs!resultado) Then
            UCCURSADAS = rs!resultado
        End If
    Else
        UCCURSADAS = 0
    End If
End Function

Function NOTAXUC(trayecto As String, codparticipante As Long) As Integer
    On Error Resume Next
    Dim rs As ADODB.Recordset
    NOTAXUC = 0
    Set rs = cn.Execute("select sum((inscripcionucurricular.calificacion * unidadcurricular.ucunidadcurricular)) as resultado from inscripcionucurricular, unidadcurricular, pnf, pensumunidadcurricular" & _
            " where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular and unidadcurricular.codunidadcurricular=pensumunidadcurricular.unidadcurricular_codunidadcurricular and" & _
            " pensumunidadcurricular.pensum_codpensum=" & codpensum & " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & FrmConstancias.txtespecialidad.Text & "' and unidadcurricular.traunidadcurricular='" & trayecto & "'" & _
            " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and substring(inscripcionucurricular.condicion from 1 for 4)='Apro'")
    If Not rs.BOF Then
        'Do While Not rs.EOF
        '    If rs!nota <> "A" Or rs!nota <> "EQ" Then
                NOTAXUC = NOTAXUC + rs!resultado
        '    End If
        '    rs.MoveNext
        'Loop
    Else
        NOTAXUC = 0
    End If
End Function

Function TOTALUCCURSADAS(codparticipante As Long, inicio As Integer, fin As Integer) As Integer
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    TOTALUCCURSADAS = 0
    If inicio = 1 Then
        sql = "select regimen as resultado from pnf,regimen where regimen.codregimen=pnf.regimen_codregimen and pnf.despnf='" & Form8.txtespecialidad.Text & "'"
        If funciones.ExisteRegistro(sql, cn) = True Then 'en caso de pnf
            Set rs = cn.Execute("select sum(unidadcurricular.ucunidadcurricular) as resultado from inscripcionucurricular, unidadcurricular, pnf Where inscripcionucurricular.unidadcurricular_codunidadcurricular = unidadcurricular.codunidadcurricular" & _
                                " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & "and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & _
                                " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and substring (inscripcionucurricular.condicion from 1 for 4)='Apro'")
        Else ' en caso de semestre o año
            Set rs = cn.Execute("select sum(unidadcurricular.ucunidadcurricular) as resultado from inscripcionucurricular, unidadcurricular, pnf where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular" & _
                        " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & " and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and" & _
                        " substring (inscripcionucurricular.condicion from 1 for 4)='Apro'")
        End If
    Else    'transicion
        Set rs = cn.Execute("select sum(unidadcurricular.ucunidadcurricular) as resultado from inscripcionucurricular, unidadcurricular, pnf where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular" & _
                        " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & " and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and" & _
                        " substring (inscripcionucurricular.condicion from 1 for 4)='Apro'")
    End If
    If Not rs.BOF Then
        TOTALUCCURSADAS = rs!resultado
    Else
        TOTALUCCURSADAS = 0
    End If
End Function

Function TOTALNOTAXUC(codparticipante As Long, inicio As Integer, fin As Integer) As Integer
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim sql As String
    
    TOTALNOTAXUC = 0
    If inicio = 1 Then  'en caso de pnf
        sql = "select regimen as resultado from pnf where despnf='" & Form8.txtespecialidad.Text & "'"
        If funciones.CampoBooleano(sql, cn) = True Then
            Set rs = cn.Execute("select sum((inscripcionucurricular.calificacion * unidadcurricular.ucunidadcurricular)) as resultado from inscripcionucurricular, unidadcurricular, pnf where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular" & _
                        " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & " and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and" & _
                        " unidadcurricular.traunidadcurricular<>'0" & fin & "' and unidadcurricular.ucunidadcurricular<>0")
        Else    'en caso de semestre o año
            Set rs = cn.Execute("select sum((inscripcionucurricular.calificacion * unidadcurricular.ucunidadcurricular)) as resultado from inscripcionucurricular, unidadcurricular, pnf where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular" & _
                        " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & " and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and" & _
                        " substring(inscripcionucurricular.condicion from 1 for 4)='Apro' and unidadcurricular.ucunidadcurricular<>0")
        End If
    Else
        Set rs = cn.Execute("select sum((inscripcionucurricular.calificacion * unidadcurricular.ucunidadcurricular)) as resultado from inscripcionucurricular, unidadcurricular, pnf where inscripcionucurricular.unidadcurricular_codunidadcurricular=unidadcurricular.codunidadcurricular" & _
                        " and unidadcurricular.pnf_codpnf=pnf.codpnf and pnf.despnf='" & Form9.txtespecialidad.Text & "' and cast(unidadcurricular.traunidadcurricular as int)>=" & inicio & " and cast(unidadcurricular.traunidadcurricular as int)<=" & fin & " and inscripcionucurricular.participantes_codparticipantes=" & codparticipante & " and" & _
                        " unidadcurricular.traunidadcurricular<>'0" & fin & "' and unidadcurricular.ucunidadcurricular<>0")
    End If
    If Not rs.BOF Then
        'Do While Not rs.EOF
        '    If rs!nota_mat <> "A" Or rs!nota_mat <> "EQ" Then
                TOTALNOTAXUC = rs!resultado
        '    End If
        '    rs.MoveNext
        'Loop
    Else
        TOTALNOTAXUC = 0
    End If
End Function

Function PARTEENTERA(cantidad As String) As String
    lugardec = InStr(1, cantidad, ",")
    If lugardec = 0 Then
       PARTEENTERA = CInt(cantidad)
    Else
       PARTEENTERA = Mid(cantidad, 1, lugardec - 1)
    End If
End Function

Function PARTEDECIMAL(cantidad As String) As String
    lugardec = InStr(1, cantidad, ",")
    If lugardec = 0 Then
        PARTEDECIMAL = 0
    Else
        PARTEDECIMAL = Mid(cantidad, lugardec + 1, 2)
    End If
End Function

Sub HOJA_DE_RESUMEN(codparticipante As Long, codpromocion As Integer, pag As Integer, cedula As String, participante As String, control As String)
    Dim sql As String
    Dim codpensum As Integer
    Dim codcohorte As Integer
    Dim rs As ADODB.Recordset
    Dim titulo As String
    Dim ent As String
    Dim dec As String
    Dim iag As String
    Dim indicepromo As Double
    
    funciones.Conectar
    sql = "select codpensum as resultado from pensum, pnf where pensum.pnf_codpnf=pnf.codpnf and pnf.despnf='" & FrmConstancias.txtespecialidad.Text & "' and pensum.despensum='" & FrmConstancias.txtpensum.Text & "'"
    codpensum = funciones.CampoEntero(sql, cn)
    sql = "select codcohorte as resultado from cohorte where descohorte='" & FrmConstancias.txtcohorte.Text & "' and pensum_codpensum=" & codpensum
    codcohorte = funciones.CampoEntero(sql, cn)
    'sql = "select promocion_codpromocion as resultado from cohortepromocion where cohorte_codcohorte=" & codcohorte
    'codpromocion = funciones.CampoEntero(sql, cn)
    'vectorimpresion(indice) = "RESUMEN" & res
    'indice = indice + 1
    'AppExcel.Application.Sheets("RESUMEN" & res).Select
    AppExcel.Cells(pag, 9).Value = control
    AppExcel.Cells(pag + 9, 2).Value = cedula
    AppExcel.Cells(pag + 9, 5).Value = participante
    
    'titulo
    sql = "select titulos_codtitulos as resultado from promocion where codpromocion=" & codpromocion
    Select Case funciones.CampoEntero(sql, cn)
        Case 2
            titulo = "TÉCNICO SUPERIOR UNIVERSITARIO EN "
        Case 3
            titulo = "INGENIERO EN "
        Case 4
            titulo = "LICENCIADO EN "
    End Select
    sql = "select titpromocion as resultado from promocion where codpromocion=" & codpromocion
    titulo = titulo & funciones.CampoString(sql, cn)
    AppExcel.Application.Cells(pag + 12, 1).Value = titulo
    'fecha
    sql = "select fecpromocion as resultado from promocion where codpromocion=" & codpromocion
    AppExcel.Application.Cells(pag + 15, 2).Value = funciones.fechacompleta(CampoString(sql, cn))
    'libro
    sql = "select liparticipantepromocion as resultado from participantepromocion where participantes_codparticipantes=" & codparticipante & " and promocion_codpromocion=" & codpromocion
    AppExcel.Application.Cells(pag + 15, 7).Value = AppExcel.Application.Cells(13, 4).Value & " " & funciones.CampoString(sql, cn)
    'folio
    sql = "select foparticipantepromocion as resultado from participantepromocion where participantes_codparticipantes=" & codparticipante & " and promocion_codpromocion=" & codpromocion
    AppExcel.Application.Cells(pag + 15, 9).Value = funciones.formatofolio(funciones.CampoString(sql, cn))
    'indice participante
    sql = "select iaparticipantepromocion as resultado from participantepromocion where participantes_codparticipantes=" & codparticipante & " and promocion_codpromocion=" & codpromocion
    iag = funciones.CampoDouble(sql, cn)
    AppExcel.Application.Cells(pag + 20, 4).Value = ConvertirNumerosEnLetras(CStr(iag))
    'indice promocion
    sql = "select indpromocion as resultado from promocion where codpromocion=" & codpromocion
    indicepromo = funciones.CampoDouble(sql, cn)
    AppExcel.Application.Cells(pag + 22, 4).Value = ConvertirNumerosEnLetras(CStr(indicepromo))
    'integrantes promocion
    sql = "select intpromocion as resultado from promocion where codpromocion=" & codpromocion
    AppExcel.Application.Cells(pag + 24, 4).Value = ConvertirNumerosEnLetras(CStr(funciones.CampoEntero(sql, cn)))
    'lugar participante
    sql = "select pgparticipantepromocion as resultado from participantepromocion where participantes_codparticipantes=" & codparticipante & " and promocion_codpromocion=" & codpromocion
    AppExcel.Application.Cells(pag + 26, 4).Value = ConvertirNumerosEnLetras(CStr(funciones.CampoEntero(sql, cn)))
    cn.Close
End Sub

Sub HOJA_DE_RESUMEN_Manual(cedula As String, participante As String, control As String, fecha As String, integrantes As Integer, puesto As Integer, tit As Integer, indicepromo As Double, libro As Integer, folio As Integer, pag As Integer)
    Dim ent As Integer
    Dim dec As Integer
    Dim especialidad As String
    Dim iag As Double
    
    'AppExcel.Application.Sheets("RESUMEN").Select
    AppExcel.Cells(pag, 9).Value = control
    AppExcel.Cells(pag + 9, 2).Value = cedula
    AppExcel.Cells(pag + 9, 5).Value = participante
    'titulo
    Select Case tit
        Case 1
            titulo = "TÉCNICO SUPERIOR UNIVERSITARIO EN "
        Case 2
            titulo = "INGENIERO EN "
        Case 3
            titulo = "LICENCIADO EN "
    End Select
    If InStr(1, FrmConstancias.txtespecialidad.Text, "(") > 0 Then
        especialidad = Mid(FrmConstancias.txtespecialidad.Text, 1, InStr(1, FrmConstancias.txtespecialidad.Text, "(") - 1)
    Else
        especialidad = FrmConstancias.txtespecialidad.Text
    End If
    If especialidad = "MATERIALES" Or especialidad = "MATERIALES IND." Then
        If tit = 1 Then
            especialidad = "POLIMEROS"
        Else
            especialidad = "MATERIALES INDUSTRIALES"
        End If
    End If
    AppExcel.Application.Cells(pag + 12, 1).Value = titulo & " " & especialidad
    'fecha
    AppExcel.Application.Cells(pag + 15, 2).Value = funciones.fechacompleta(fecha)
    'libro
    AppExcel.Application.Cells(pag + 15, 7).Value = libro
    'folio
    AppExcel.Application.Cells(pag + 15, 9).Value = folio
    'indice participante
    iag = Round(totalnotauc / totaluccur, 2)
    AppExcel.Application.Cells(pag + 20, 4).Value = ConvertirNumerosEnLetras(CStr(iag))
    'indice promocion
    AppExcel.Application.Cells(pag + 22, 4).Value = ConvertirNumerosEnLetras(CStr(indicepromo))
    'integrantes promocion
    AppExcel.Application.Cells(pag + 24, 4).Value = ConvertirNumerosEnLetras(CStr(integrantes))
    'lugar participante
    AppExcel.Application.Cells(pag + 26, 4).Value = ConvertirNumerosEnLetras(CStr(puesto))
    cn.Close
End Sub

Function numerotrayecto(trayecto As String, codpensum As Integer, codpnf As Integer, cohorte As String) As String
    Dim sql As String
    
    sql = "select pensumnivel.trapensumnivel as resultado from cohorte,pensum, pensumnivel, nivel,pnf where cohorte.nivel_codnivel=nivel.codnivel and cohorte.pnf_codpnf=pnf.codpnf and" & _
          " nivel.codnivel=pensumnivel.nivel_codnivel and pensumnivel.pensum_codpensum=pensum.codpensum and pnf.codpnf=" & codpnf & " and cohorte.descohorte='" & cohorte & "' and pensum.codpensum=" & codpensum
    
    If trayecto = "Inicial" Or trayecto = "Transición" Then
        numerotrayecto = "0" & funciones.CampoEntero(sql, cn)
    Else
        numerotrayecto = trayecto
    End If
End Function

Function nombremes(mes As Integer) As String
    Select Case mes
        Case 1: nombremes = "Enero"
        Case 2: nombremes = "Febrero"
        Case 3: nombremes = "Marzo"
        Case 4: nombremes = "Abril"
        Case 5: nombremes = "Mayo"
        Case 6: nombremes = "Junio"
        Case 7: nombremes = "Julio"
        Case 8: nombremes = "Agosto"
        Case 9: nombremes = "Septiembre"
        Case 10: nombremes = "Octubre"
        Case 11: nombremes = "Noviembre"
        Case 12: nombremes = "Diciembre"
    End Select
End Function

'Devuelve el último días del Mes
Public Function fin_del_Mes(fecha As Variant) As String
  
    If IsDate(fecha) Then
        fin_del_Mes = DateAdd("m", 1, fecha)
        fin_del_Mes = DateSerial(Year(fin_del_Mes), Month(fin_del_Mes), 1)
        fin_del_Mes = DateAdd("d", -1, fin_del_Mes)
        fin_del_Mes = Mid(fin_del_Mes, 7, 4) & "-" & Mid(fin_del_Mes, 4, 2) & "-" & Mid(fin_del_Mes, 1, 2)
    End If
  
End Function

Public Sub RegistrardatosPromocionManual(ByRef fechagrado As String, ByRef integrantes As Integer, ByRef puesto As Integer, ByRef tit As Integer, ByRef indicepromo As String, ByRef libro As Integer, ByRef folio As Integer)
    
    Do
        fechagrado = InputBox("Fecha dd/mm/aaaa:", "Fecha del Acto de grado")
        If Len(fechagrado) <> 10 Then
            MsgBox "debe indicar la fecha en el formato solicitado", vbCritical
        End If
    Loop Until Len(fechagrado) >= 10
    Do
        integrantes = InputBox("Cantidad:", "Número de Integrantes de la Promoción")
        If integrantes <= 0 Then
            MsgBox "debe indicar la cantidad de integrantes de la promoción", vbCritical
        End If
    Loop Until integrantes > 0
    Do
        puesto = InputBox("Lugar:", "Puesto dentro de la Promoción")
        If puesto <= 0 Then
            MsgBox "debe indicar el puesto ocupado en la promoción", vbCritical
        End If
    Loop Until integrantes > 0
    Do
        indicepromo = InputBox("indice promoción:", "Indice académico de la promoción")
        If Len(indicepromo) <= 0 Then
            MsgBox "debe indicar el indice académico de la promoción", vbCritical
            indicepromo = 0
        Else
            indicepromo = Replace(indicepromo, ".", ",")
            If CDbl(indicepromo) <= 0 Then
                MsgBox "el indice académico de la promoción no puede ser negativo o cero", vbCritical
                indicepromo = 0
            End If
        End If
    Loop Until Len(indicepromo) > 0 And CDbl(indicepromo) > 0
    Do
        libro = InputBox("Libro:", "Libro de grado")
        If libro <= 0 Then
            MsgBox "debe indicar el libro del graduado", vbCritical
        End If
    Loop Until libro > 0
    Do
        folio = InputBox("Folio:", "Folio de grado")
        If folio <= 0 Then
            MsgBox "debe indicar el folio del graduado", vbCritical
        End If
    Loop Until folio > 0
    Do
        tit = InputBox("1.-T.S.U" & Chr(13) & "2.-Ingeniero" & Chr(13) & "3.-Licenciado", "Titulo")
        If tit < 1 Or tit > 3 Then
            MsgBox "debe seleccionar una opcion valida", vbCritical
        End If
    Loop Until tit >= 1 And tit <= 3
End Sub

Public Sub Encabezado_Hoja(cedula As String, participante As String, especialidad As String, control As String, pagina As Integer, fila As Integer, pag As Integer)
    
    If pag > 1 Then
        AppExcel.Application.Sheets(1).Select
        AppExcel.Application.Range("A1:J5").Select
        AppExcel.Application.Selection.Copy
        AppExcel.Application.Sheets("CERTIFICADAS1").Select
        AppExcel.Application.Range("A" & pag).Select
        AppExcel.Application.ActiveSheet.Paste
    End If
    AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila, 3).Value = cedula
    AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila + 1, 3).Value = participante
    If InStr(1, especialidad, "(") > 0 Then
        AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila + 2, 3).Value = Mid(especialidad, 1, InStr(especialidad, " (") - 1)
    Else
        AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila + 2, 3).Value = especialidad
    End If
    'coloca el numero de pagina
    AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila + 3, 1).Value = AppExcel.Sheets("CERTIFICADAS1").Cells(pag + fila + 3, 1).Value & " " & pagina
    If pag = 0 Then
        AppExcel.Sheets("CERTIFICADAS1").Cells(1, 9).Value = control
    Else
        AppExcel.Sheets("CERTIFICADAS1").Cells(pag, 9).Value = control
    End If
End Sub

Public Sub Cuerpo_Hoja(cedula As String, sql As String, ByRef pag As Integer, ByRef pagina As Integer, ByRef fila As Integer, control As String, codparticipante As Long, codpromocion As Integer, participante As String, especialidad As String, regpromocion As Integer)
    Dim rs1 As Recordset
    
    Dim fechagrado As String
    Dim integrantes As Integer
    Dim puesto As Integer
    Dim tit As Integer
    Dim indicepromo As String
    Dim libro As Integer
    Dim folio As Integer
    Dim sql2 As String
    
    funciones.Conectar
    Set rs1 = cn.Execute(sql)
    If Not rs1.BOF Then
        Do While Not rs1.EOF
            With AppExcel.Application
            
                funciones.ENCABEZADOSEMESTRE pag + fila - 1, rs1!resultado, rs1!resultado2
                fila = fila + 2
                'TRANSFIERE LAS NOTAS DEL PERIODO A LA PLANTILLA
                'fila = pag + fila - 1
                
                funciones.COLOCARNOTA rs1!resultado, funciones.codparticipante, pag, pagina, fila, control
                'COLOCA LA LEYENDA PARA IRA DEL SEMESTRE
                If fila + 2 > 59 Then
                    pagina = pagina + 1
                    pag = pag + 59
                    fila = 1
                    funciones.Encabezado_Hoja cedula, participante, especialidad, control, pagina, fila, pag
                    fila = fila + 6
                End If
                fila = fila + 1
                funciones.PIESEMESTRE pag + fila - 1
                funciones.RESUMENPERIODO rs1!resultado, codparticipante, pag, pag + fila - 1
                fila = fila + 3
                If (fila) > 59 Then
                    pagina = pagina + 1
                    pag = pag + 59
                    fila = 1
                    funciones.Encabezado_Hoja cedula, participante, especialidad, control, pagina, fila, pag
                    fila = fila + 5
                End If
            End With
            rs1.MoveNext
        Loop
        If (fila + 10) > 59 Then
            pagina = pagina + 1
            pag = pag + 59
            fila = 1
            funciones.Encabezado_Hoja cedula, participante, especialidad, control, pagina, fila, pag
            fila = fila + 5
        End If
        'COLOCA LA LEYENDA PARA IRA DEL SEMESTRE
        funciones.PIEDOCUMENTO pag + fila
        funciones.RESUMENTOTAL codparticipante, pag + fila
        fila = fila + 9
        
        If (fila + 2) > 59 Then
            pagina = pagina + 1
            pag = pag + 59
            fila = 1
            funciones.Encabezado_Hoja cedula, participante, especialidad, control, pagina, fila, pag
            fila = fila + 5
        End If
        funciones.Fecha_Solicitud pag + fila
        pagina = pagina + 1
        pag = pag + 59
        fila = 1
        'aqui coloca los datos en la hoja de resumen de la graduacion
        If regpromocion = 1 Then    'si esta registrado en la promocion
            funciones.Copiar_Hoja_Resumen pag
            funciones.HOJA_DE_RESUMEN codparticipante, codpromocion, pag, cedula, participante, control
            pag = pag + 59
        Else  'si se cargan los datos de la promocion de manera manual
            funciones.RegistrardatosPromocionManual fechagrado, integrantes, puesto, tit, indicepromo, libro, folio
            funciones.Copiar_Hoja_Resumen pag
            funciones.HOJA_DE_RESUMEN_Manual cedula, participante, control, fechagrado, integrantes, puesto, tit, CDbl(indicepromo), libro, folio, pag
            pag = pag + 59
        End If
    End If
End Sub

Sub Encabezado_listado_promocion(fil As Integer, pag As Integer)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A1:F10").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets(pag).Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
End Sub

'Sub Colocar_Centrado(valor As String, fila As Integer, columna As Integer, alineacion As Integer)
'    Dim rango As String
'
'    AppExcel.Application.Cells(fila, columna).Value = valor
'    rango = Celda(columna) & fila
'    AppExcel.Application.Range(rango).Select
'    With AppExcel.Application.Selection
'        .HorizontalAlignment = alineacion
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'End Sub

'Function Celda(col As Integer) As String
'    Select Case col
'        Case 1: Celda = "A"
'        Case 2: Celda = "B"
'        Case 3: Celda = "C"
'        Case 4: Celda = "D"
'        Case 5: Celda = "E"
'        Case 6: Celda = "F"
'        Case 7: Celda = "G"
'        Case 8: Celda = "H"
'        Case 9: Celda = "I"
'        Case 10: Celda = "J"
'    End Select
'End Function

Public Sub RegistroEvento(cedula As String, fecha As String, evento As String, usuario As String, cedparticipante As String, participante As String)
  Dim sql As String
  Dim codevento As Long
  
  funciones.Conectar2
    sql = "select max(codevento) as resultado from eventos"
    codevento = funciones.proximocodigoregistro(sql, cn2)
    sql = "insert into eventos values(" & codevento & ",'" & cedula & "','" & fecha & "','" & evento & "','" & Time() & "','" & usuario & "','" & cedparticipante & "','" & participante & "')"
    cn2.Execute (sql)
  cn2.Close
End Sub

Public Function FormatoFechaConsulta(fecha As String) As String
  Dim dd As String      'dia
  Dim mm As String      'mes
  Dim aa As String      'aÃ±o
  Dim res As String
  dd = Mid(fecha, 1, 2)
  mm = Mid(fecha, 4, 2)
  aa = Mid(fecha, 7, 4)
  FormatoFechaConsulta = dd & "-" & mm & "-" & aa
End Function

Sub Copiar_Hoja(fil As Integer, rango As String, pagina1 As String, pagina2 As String)
    AppExcel.Application.Sheets(pagina2).Select
    AppExcel.Application.Range(rango).Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets(pagina1).Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
End Sub

Sub Copiar_Hoja_Resumen(fila As Integer)
    AppExcel.Application.Sheets("RESUMEN").Select
    AppExcel.Application.Rows("1:59").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fila).Select
    AppExcel.Application.ActiveSheet.Paste
End Sub

Sub Fecha_Solicitud(fil As Integer)
    AppExcel.Application.Sheets(1).Select
    AppExcel.Application.Range("A21:a23").Select
    AppExcel.Application.Selection.Copy
    AppExcel.Application.Sheets("CERTIFICADAS1").Select
    AppExcel.Application.Range("A" & fil).Select
    AppExcel.Application.ActiveSheet.Paste
    AppExcel.Sheets("CERTIFICADAS1").Cells(fil, 1).Value = "A petición de la parte interesada para los fines que a la misma convegan, se extiende la presente en la ciudad de Valencia, Carabobo a los " & Day(Date) & " dias de mes de " & funciones.nombremes(Month(Date)) & "·del año " & Year(Date)
End Sub

Function ConvertirNumerosEnLetras(cantidad As String) As String
    ent = funciones.PARTEENTERA(CStr(cantidad))
    dec = funciones.PARTEDECIMAL(CStr(cantidad))
    If dec = "0" Then   'sin decimales
        ConvertirNumerosEnLetras = ent & "  (" & Trim(funciones.EnLetras(CStr(ent))) & ")"
    Else
        If InStr(1, cantidad, ",0") > 0 Then
            cadenadecimal = funciones.EnLetras(0) & " " & funciones.EnLetras(CInt(dec))
            ConvertirNumerosEnLetras = ent & "," & dec & "  (" & funciones.EnLetras(CStr(ent)) & " con " & Trim(cadenadecimal) & ")"
        Else
            If Len(Trim(dec)) = 1 Then
                dec = dec & "0"
                cadenadecimal = funciones.EnLetras(CInt(dec))
            Else
                cadenadecimal = funciones.EnLetras(CInt(dec))
            End If
            ConvertirNumerosEnLetras = ent & "," & dec & "  (" & funciones.EnLetras(CStr(ent)) & " con " & Trim(cadenadecimal) & ")"
        End If
    End If
End Function

Public Function EnLetras(numero As String) As String
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String

    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "," Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
    flag = "N"
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "0"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "cero "
                            'Else
                            '   expresion = expresion & "un "
                            End If
                        End If
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                               expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                   'aa = Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 6
                    expresion = expresion & "mil "
                Else
                    expresion = expresion & "mil "
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
                Else
                    expresion = expresion & "millones "
                End If
                flag = "N"
            End If
        Next paso
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci ' & "/100"
            Else
                EnLetras = expresion & "con " & deci ' & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
End Function

