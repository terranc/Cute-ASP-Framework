<%
'**********
'	class		: A database class
'	File Name	: db.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-4
'**********


'**********
'	ʾ��
'**********

'**********
'	������
'**********
Class Class_Db
    Private ConnStr, SqlLocalPath, rs
    Public ServerIp		'���ݿ�����������
	Public ConnectionType	'���ݿ��������� -- 1.ACCESS 2.MSSQL
	Public Database		'���ݿ��� 
	Public Username			'�û���
	Public Password			'����
    Public Conn

    Private Sub Class_Initialize()
		ServerIp			= "(local)"
		ConnectionType		= "MSSQL"
		Username			= "sa"
		Password			= ""
    End Sub
	
    Private Sub Class_Terminate()
        Me.Close()
		Me.CloseRs null
    End Sub
	
    '**********
    ' ������: Open
    ' ��  ��: �����ݿ�����
    '**********
    Sub Open()
		On Error Resume Next
		If Me.ConnectionType = "ACCESS" Then
			SqlLocalPath = Replace(Request.ServerVariables("PATH_TRANSLATED"),Replace(Request.ServerVariables("PATH_INFO"),"/","\"),"")
			ConnStr = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & SqlLocalPath & Database
		ElseIf Me.ConnectionType = "MSSQL" Then
			ConnStr = "Provider=SQLOLEDB.1;DATA SOURCE=" & Me.ServerIp & ";UID="&Username&";PWD="&Password&";Database="&Database&";Pooling=true; MAX Pool Size=512;Min Pool Size=50;Connection Lifetime=30"
		End If
		Set Conn = CreateObject("ADODB.Connection")
		Conn.Open ConnStr
		If Err Then
			Set Conn = Nothing
			Response.Write Err.Description
			Err.Clear
			Response.End()
		End If
		On Error GoTo 0
    End Sub

    '**********
    ' ������: Close
    ' ��  ��: �رղ��ͷ����ݿ�����
    '**********
	Sub Close()
        On Error Resume Next
		Me.Conn.Close
		Set Me.Conn = Nothing
		Err.Clear
		On Error Goto 0
	End Sub
	
    '**********
    ' ������: CloseRs
    ' ��  ��: OutRs as recordset object
    ' ��  ��: �رղ��ͷ�ָ������
    '**********
	Sub CloseRs(OutRs)
		If IsObject(OutRs) Then
			On Error Resume Next
			OutRs.close
			Set OutRs = Nothing
			On Error Goto 0
		End If
	End Sub

    Sub SetRs(OutRs, strsql, CursorAndLockType)
		Set OutRs = Server.Createobject("ADODB.Recordset")
        OutRs.Open strsql, Conn, 1, CursorAndLockType
    End Sub

    '**********
    ' ������: Exec
    ' ��  ��: OutRs as recordset object
    ' ��  ��: sql as sql string
    ' ��  ��: ִ��sql��䣬���ǲ�ѯ����򽫽�����ظ�OutRs
    '**********
    Sub Exec(OutRs, sql)
        If InStr(UCase(sql), UCase("select"))>0 Then
            Set OutRs = Conn.Execute(sql)
        Else
            Call Conn.Execute(sql)
            OutRs = 1
        End If
    End Sub

	
    '**********
    ' ������: Close
    ' ��  ��: ��ʼ����
    '**********
	Sub BeginTrans()
		Me.Conn.BeginTrans()
	End Sub
	
    '**********
    ' ������: Close
    ' ��  ��: �ع�����
    '**********
	Sub RollBackTrans()
		Me.Conn.RollBackTrans()
	End Sub

    '**********
    ' ������: Close
    ' ��  ��: �ύ����
    '**********
	Sub CommitTrans()
		Me.Conn.CommitTrans()
	End Sub

    '**********
    ' ������: GetRecordObject
    ' ��  ��: source as sql string or recordset object
    ' ��  ��: ��һ����¼תΪһ���ֵ����
    '**********
	Sub GetRowObject(obj,source)
		Dim i, rs
		If TypeName(source) = "Recordset" Then
			Set rs = source
		Else
			Me.Exec rs,source
		End If
		If rs.eof Then Exit Sub
		Set obj = CreateObject("Scripting.Dictionary")
		obj.CompareMode = 1
		For i=0 To rs.fields.count-1
			obj.Add rs.Fields(i).Name,rs(i).Value
		Next
		Me.closeRs rs
	End Sub

    '**********
    ' ������: Insert
    ' ��  ��: Table as Data Table
    ' ��  ��: Params as Dictionary
    ' ��  ��: �����¼
    '**********
	Function Insert(Table,Params)
		Dim sqlCmd, sqlCmd_a, sqlCmd_b, parameteres, oParams
		Dim iName
		sqlCmd = "Set nocount on" & vbCrlf
		sqlCmd = sqlCmd & "Insert Into "&Table&" ("
		parameteres = " "
		Set oParams = CreateObject("Scripting.Dictionary")
		If Not IsNull(params) Then
			For Each iName in params
				sqlCmd_a = sqlCmd_a & iName & ","
				sqlCmd_b = sqlCmd_b & "@" & iName & ","
				parameteres = parameteres & "@" & iName & " varchar(8000)" & ","
			Next
		End If
		sqlCmd_a = Left(sqlCmd_a,Len(sqlCmd_a)-1)
		sqlCmd_b = Left(sqlCmd_b,Len(sqlCmd_b)-1)
		sqlCmd = sqlCmd & sqlCmd_a & ")  values(" & sqlCmd_b & ")"
		sqlCmd = sqlCmd & vbCrlf & "select Cast(IsNull(SCOPE_IDENTITY(),-100) as int)"
		parameteres = Left(parameteres,Len(parameteres)-1)
		oParams.Add "@stmt",sqlCmd
		oParams.Add "@parameters",parameteres
		If Not IsNull(Params) Then
			For Each iName in Params
				oParams.Add "@"&iName,Params(iName)&""
			Next
		End If
		Insert = Me.ExecuteScalar("sp_executesql",oParams)
		Set oParams = Nothing
	End Function

    '**********
    ' ������: Update
    ' ��  ��: Table as Data Table
    ' ��  ��: Params as Dictionary
    ' ��  ��: Where as �������
    ' ��  ��: ���¼�¼
    '**********
	Function Update(Table,Params,Where)
		Dim sqlCmd, parameteres, oParams
		Dim iName
		sqlCmd = "Set nocount on" & vbCrlf
		sqlCmd = sqlCmd & "Update "&Table&" set "
		parameteres = " "
		Set oParams = CreateObject("Scripting.Dictionary")
		If Not IsNull(params) Then
			For Each iName in params
				If InStr(iName,"#") > 0 Then
					params.Key(iName) = Replace(iName,"#","")
					iName = Replace(iName,"#","")
					sqlCmd = sqlCmd & iName & "=" & iName & " + @" & iName & ","
				Else
					sqlCmd = sqlCmd & iName & "=@" & iName & ","
				End If
				parameteres = parameteres & "@" & iName & " varchar(8000)" & ","
			Next
		End If
		sqlCmd = Left(sqlCmd,Len(sqlCmd)-1)
		If Trim(Where) <> "" Then sqlCmd=sqlCmd&" Where "&Where&""
		sqlCmd = sqlCmd & vbCrlf & "select CAST(IsNull(@@ROWCOUNT,-100) as int)"
		parameteres = Left(parameteres,Len(parameteres)-1)
		oParams.Add "@stmt",sqlCmd
		oParams.Add "@parameters",parameteres
		If Not IsNull(Params) Then
			For Each iName in Params
				oParams.Add "@"&iName,Params(iName)&""
			Next
		End If
		Update = Me.ExecuteScalar("sp_executesql",oParams)
		Set oParams = Nothing
	End Function

	'**********
	' ����	��	ִ�д洢���̲����ؼ�¼��
	'**********
	Function ExecuteRecordSet(commandName , ByVal params)
		Set ExecuteRecordSet = ExecuteSqlCommand(commandName , params , 2)
	End Function
	
	'**********
	' ����	��	ִ�д洢���̲����ؼ�¼����һ�е�һ��
	'**********
	Function ExecuteScalar(commandName , ByVal params)
		Set rs = ExecuteSqlCommand(commandName , params , 2)
		If Not rs.EOF And Not rs.BOF Then
			ExecuteScalar = rs(0).Value
		Else
			ExecuteScalar = NULL
		End If
		rs.Close
		Set rs = Nothing
	End Function
	
	'**********
	' ����	��	ִ�д洢���̲�����һ��ֵ
	'**********
	Function ExecuteReturnValue(commandName , ByVal params)
		ExecuteReturnValue = ExecuteSqlCommand(commandName , params , 1)
	End Function
	
	'**********
	' ����	��	ִ�д洢���̲������κ�����
	'**********
	Function ExecuteNonQuery(commandName , ByVal params)
		ExecuteNonQuery = ExecuteSqlCommand(commandName , params , 0)
	End Function
	
	'**********
	' ����	��	ִ�д洢���̲����ؼ�¼����һ��ֵ
	'**********
	Function ExecuteRecordsetAndValue(commandName , ByVal params)
		ExecuteRecordsetAndValue = ExecuteSqlCommand(commandName , params , 3)
	End Function
	
	'**********
	' ����	��	ִ�д洢����
	'**********
	' commandName		�洢��������
	' params			�������ϣ�����ʹ�� Scripting.Dictionary ������
	' returnMode		����ģʽ
	'					0	�������κβ��������
	'					1	ִ�к�õ�����ֵ
	'					2	ִ�к�õ���¼��
	'					3	ִ�к�õ�����ֵ�ͼ�¼��
	'**********
	Private Function ExecuteSqlCommand(commandName , ByVal params , returnMode)
		Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
		Dim iName : iName = ""
		Dim RSReturn : Set RSReturn = Nothing
		DIM RSStream	: SET RSStream	= Server.CreateObject("ADODB.Stream")
		Dim ReturnValue : ReturnValue = ""
		
		cmd.ActiveConnection = Me.conn
		cmd.CommandText = commandName
		cmd.CommandType = 4
		cmd.NamedParameters = True
		cmd.Prepared = True
	
		If returnMode = 1 Or returnMode = 3 Then
			cmd.Parameters.Append cmd.CreateParameter("@ReturnValue" , 2 , 4)
		End If
		If Not IsNull(params) Then
			For Each iName in params
				If iName <> "@stmt" And iName <> "@statement" And iName <> "@parameters" Then
					If Len(params(iName)) < 4000 Or IsNumeric(params(iName)) Then
						cmd.Parameters.Append cmd.CreateParameter(iName , 202 , 1 , 4000 , params(iName)&"")
					Else
						cmd.Parameters.Append cmd.CreateParameter(iName , 203 , 1 , Len(params(iName)) + 2 , params(iName)&"")
					End If
				Else
					cmd.Parameters.Append cmd.CreateParameter(iName , 202 , 1 , 4000 , params(iName))
				End If
			Next
		End If
	
		Select Case returnMode
			' ִ�к�õ�����ֵ
			Case 1
				Call cmd.Execute(, , 128)
				ExecuteSqlCommand = cmd("@ReturnValue").Value
				' ִ�к�õ���¼��
			Case 2
				Set ExecuteSqlCommand = cmd.Execute()
			Case 3
				Set RSReturn = cmd.Execute()
				Call RSReturn.Save(RSStream,1)
				RSReturn.Close
				Call RSReturn.Open(RSStream)

				ExecuteSqlCommand = Array(RSReturn, cmd("@ReturnValue").Value)
				' Ĭ�Ϸ�ʽ���������κβ��������
			Case Else
				Call cmd.Execute(ExecuteSqlCommand, , 128)
		End Select
	
		Set cmd = Nothing
	End Function
End Class

%>