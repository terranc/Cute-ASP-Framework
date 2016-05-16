<%
'**********
'	class		: some function Library
'	File Name	: function.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-3
'**********

'**********
'函数名：ShowErr
'参  数：message	-- 错误信息
'作  用：显示错误信息
'**********
Sub ShowErr(message)
    die "<html><head><title>Exception page</title><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><style type=""text/css""><!--" & vbCrlf & "* { margin:0; padding:0 }" & vbCrlf & "body { background:#333; color:#0f0; font:14px/1.6em ""宋体"", Verdana, Arial, Helvetica, sans-serif; }" & vbCrlf & "dl { margin:20px 40px; padding:20px; border:3px solid #f63; }" & vbCrlf & "dt { margin:0 0 0.8em 0; font-weight:bold; font-size:1.6em; }" & vbCrlf & "dd { margin-left:2em; margin-top:0.2em; }" & vbCrlf & "--></style></head><body><div id=""container""><dl><dt>Description:</dt><dd><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Position:</span> " & message & "</dd></dl></div></body></html>"
End Sub

'**********
'函数名：ShowException
'作  用：显示异常信息
'**********
Sub ShowException()
    echo "<p><span style=""color:#ff0;font-weight:bold;font-size:1.2em;"">Error:</span> " & Err.Number & " " & Err.Description & "</p>"
    Err.Clear
    die("")
End Sub

'**********
' 函数名: CheckPostSource
' 作  用: 检验来源地址
'**********
Function CheckPostSource()
	Dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(server_v1,8,Len(server_v2))=server_v2 Then
		CheckPostSource=True
	Else
		CheckPostSource=False
	End If
End Function

'**********
'函数名：GetSystem
'作  用：获取客户端操作系统版本
'返回值：操作系统版本名称
'**********
Function GetSystem()
	Dim System
	System = Request.ServerVariables("HTTP_USER_AGENT")
	If Instr(System,"Windows NT 5.2") Then
		System = "Win2003"
	ElseIf Instr(System,"Windows NT 5.0") Then
		System="Win2000"
	ElseIf Instr(System,"Windows NT 5.1") Then
		System = "WinXP"
	ElseIf Instr(System,"Windows NT") Then
		System = "WinNT"
	ElseIf Instr(System,"Windows 9") Then
		System = "Win9x"
	Elseif Instr(System,"unix") Or InStr(System,"linux") Or InStr(System,"SunOS") Or InStr(System,"BSD") Then
		System = "Unix"
	ElseIf Instr(System,"Mac") Then
		System = "Mac"
	Else
		System = "Other"
	End If
	GetSystem = System
End Function

'**********
'函数名：IsInstall
'作  用：检查组件是否已经安装
'参  数：obj ----组件名
'返回值：True  ----已经安装
'		 False ----没有安装
'**********
Function IsInstall(obj)
	On Error Resume Next
	IsInstall = False
	Dim xTestObj
	Set xTestObj = Server.CreateObject(obj)
	If Err.Number = 0 Then IsInstall = True
	Set xTestObj = Nothing
	Err.Clear
	On Error Goto 0
End Function

'**********
' 函数名: NoBuffer
' 作  用: no buffer
'**********
Sub NoBuffer()
	Response.Buffer = True
	Response.Expires = 0
	Response.AddHeader "Expires",0
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","no-cache,private, post-check=0, pre-check=0, max-age=0"
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
End Sub

'**********
' 函数名: Rand
' 参  数: str as the input string
' 作  用: Generate a Random integer
'**********
Function Rand(min, max)
    Randomize
    Rand = Int((max - min + 1) * Rnd + min)
End Function

'**********
' 函数名: RandStr
' 作用: Generate a specific length Random string
'**********
Function RandStr(intLength)
	Dim strSeed,seedLength,i
	strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
	seedLength = len(strSeed)
	For i=1 to intLength
		Randomize
		RandStr = RandStr & Mid(strSeed,Round((Rnd*(seedLength-1))+1),1)
	Next
End Function

'**********
' 函数名: rq
' 参  数: Requester as the request type
' 参  数: Name as the request name
' 参  数: iType as check type
' 参  数: Default as the Default string
' 作  用: Safe filter
'**********
Function rq(Requester,Name,iType,Default)
	Dim tmp
	Select Case Requester
	Case 0
		tmp = Name
	Case 1
		tmp = Trim(Request(Name))
		tmp = HtmlEncode(tmp)
	Case 2
		tmp = Trim(Request.QueryString(Name))
		tmp = HtmlEncode(tmp)
	Case 3
		tmp = Trim(Form(Name))
		tmp = HtmlEncode(tmp)
	Case 4
		tmp = Request.Cookies(Name)
		tmp = HtmlEncode(tmp)
	End Select
	If tmp = "" Then tmp = Default
	Select Case iType
	Case 0
		If IsNumeric(tmp) = False Then
			tmp = Default
		Else
			tmp = CSng(tmp)
		End If
	Case 1
		tmp = Safe(EncodeJP(tmp))
	Case 2
		If Not IsDate(tmp) Or Len(tmp) <= 0 Then 
			tmp = CDate(Default)
		Else
			tmp = CDate(tmp)
		End If
	End Select
	rq = tmp
End function

'**********
'函数名：Form
'参  数：element ---- 控件名
'作  用：获取Form控件数据
'**********
Function Form(element)
    On Error Resume Next
    If InStr(LCase(Request.ServerVariables("Content_Type")), "multipart/Form-data") Then	'multipart/Form-data
        If IsObject(Casp.Upload) = False Then
    		If Err Then
				die("no include upload class file")
			End If
        End If
		If Casp.Upload.Error <> 0 Then
			Casp.Upload.Open
		End If
        Form = Casp.Upload.Form(element)
    Else
        Form = Request.Form(element)
    End If
    On Error GoTo 0
End Function

'**********
'函数名：sqlFilter
'作  用：过滤Sql关键字
'**********
Function Safe(str)
	Safe = str
	If str = "" Then Exit Function
	str = Replace(str, "'", "''")
	str = Replace(str,"--","&#45;&#45;")
	If IsArray(FilterWord) Then
		Dim item
		str = Casp.String.RegexpReplace(str,Join(FilterWord,"|"),"XX",false)
	End If
	Safe = str
End Function

Private Function StrRepeat(length,str)
	Dim i
	For i = 1 To length
		StrRepeat = StrRepeat & str
	Next
End Function

'**********
' 函数名: CurrentURL
' 作  用: 返回当前地址
'**********
Function CurrentURL()
	Dim port : port = LCase(Request.ServerVariables("Server_Port"))
    Dim page : page = LCase(Request.ServerVariables("Script_Name"))
    Dim query : query = LCase(Request.QueryString())
	Dim url
	If CStr(port) = "80" Then 
		url = page 
	Else
		url = ":" & port & page
	End If
	If query <> "" Then
		CurrentURL = "http://" & Request.ServerVariables("server_name") & url & "?" & query
	Else
		CurrentURL = "http://" & Request.ServerVariables("server_name") & url
	End If
End Function

'**********
' 函数名: refererURL
' 作  用: 返回来源地址
'**********
Function RefererURL()
	RefererURL = Request.ServerVariables("HTTP_REFERER")
End Function

'**********
' 函数名: GetIP
' 作  用: 获取客户端IP
'**********
Function GetIP()
	Dim Ip,Tmp
	Dim i,IsErr
	Dim ForTotal
	IsErr=False
	Ip=Request.ServerVariables("REMOTE_ADDR")
	If Len(Ip)<=0 Then Ip=Request.ServerVariables("HTTP_X_ForWARDED_For")
	If Len(Ip)>15 Then 
		IsErr=True
	Else
		Tmp=Split(Ip,".")
		If Ubound(Tmp)=3 Then 
			ForTotal = Ubound(Tmp)
			For i=0 To ForTotal
				If Len(Tmp(i))>3 Then IsErr=True
			Next
		Else
			IsErr=True
		End If
	End If
	If IsErr Then 
		GetIP="1.1.1.1"
	Else
		GetIP=Ip
	End If
End Function

'**********
' 函数名: GetSelfName
' 作  用: 获取当前访问文件名
'**********
Function GetSelfName()
    GetSelfName = Request.ServerVariables("PATH_TRANSLATED")
    GetSelfName = LCase(Mid(GetSelfName, InstrRev(GetSelfName, "\") + 1, Len(GetSelfName)))
End Function

'**********
'函数名：ReturnObj
'作  用：返回一个对象
'返回值：对象包含三个参数(Code,Message,AttachObject)
'**********
Function ReturnObj()
	On Error Resume Next
	TypeName(New AopResult)
	If Err Then
		Set ReturnObj = New ReAopResult		'重定义的AopResult
		Err.Clear
	Else
		Set ReturnObj = New AopResult
	End If
	On Error Goto 0
End Function

'**********
' 函数名: EncodeJP
' 参  数: str as the input string
' 作  用: 编码日文
'**********
Function EncodeJP(ByVal str)
	EncodeJP = str
	If str="" Then Exit Function
	Dim c1 : c1 = Array("ガ","ギ","グ","ア","ゲ","ゴ","ザ","ジ","ズ","ゼ","ゾ","ダ","ヂ","ヅ","デ","ド","バ","パ","ビ","ピ","ブ","プ","ベ","ペ","ボ","ポ","ヴ")
	Dim c2 : c2 = Array("460","462","463","450","466","468","470","472","474","476","478","480","482","485","487","489","496","497","499","500","502","503","505","506","508","509","532")
	Dim i
	For i=0 to 26
		str=Replace(str,c1(i),"&#12"&c2(i)&";")
	Next
	EncodeJP = str
End Function

'**********
' 函数名: HtmlEncode
' 参  数: str as the input string
' 作  用: filter html code
'**********
Function HtmlEncode(ByVal Str)
	If Trim(Str) = "" Or IsNull(Str) Then
		HtmlEncode = ""
	Else
		str = Replace(str, "  ", "&nbsp; ")
		str = Replace(str, """", "&quot;")
		str = Replace(str, ">", "&gt;")
		str = Replace(str, "<", "&lt;")
		HtmlEncode = Str
	End If
End Function



'**********
' 函数名: HtmlDecode
' 参  数: str as the input string
' 作  用: Decode the html tag
'**********
Function HtmlDecode(ByVal str)
	If Not IsNull(str) And str <> "" Then
		str = Replace(str, "&nbsp;", " ",1,-1,1)
		str = Replace(str, "&quot;", """",1,-1,1)
		str = Replace(str, "&gt;", ">",1,-1,1)
		str = Replace(str, "&lt;", "<",1,-1,1)
		HtmlDecode = str
	End If
End Function

'**********
' 函数名: UrlDecode
' 作  用: UrlDecode — URL decode
'**********
Function UrlDecode(ByVal vstrin)
	Dim i, strreturn, strSpecial, intasc, thischr
	strSpecial = "!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%"
	strreturn = ""
	For i = 1 To Len(vstrin)
		thischr = Mid(vstrin, i, 1)
		If thischr = "%" Then
			intasc = Eval("&h" + Mid(vstrin, i + 1, 2))
			If InStr(strSpecial, Chr(intasc))>0 Then
				strreturn = strreturn & Chr(intasc)
				i = i + 2
			Else
				intasc = Eval("&h" + Mid(vstrin, i + 1, 2) + Mid(vstrin, i + 4, 2))
				strreturn = strreturn & Chr(intasc)
				i = i + 5
			End If
		Else
			If thischr = "+" Then
				strreturn = strreturn & " "
			Else
				strreturn = strreturn & thischr
			End If
		End If
	Next
	UrlDecode = strreturn
End Function

'**********
' 函数名: SetQueryString
' 作用: 重置参数
'**********
Function SetQueryString(ByVal sQuery, ByVal Name,ByVal Value)
	Dim Obj
	If Len(sQuery) > 0 Then
		If InStr(1,sQuery,Name&"=",1) = 0 Then
			If InStr(sQuery,"=") > 0 And Right(sQuery,1) <> "&" Then
				sQuery = sQuery & "&" & Name & "=" & Value
			Else
				sQuery = sQuery & Name & "=" & Value
			End If
		Else
			Set Obj = New Regexp
			Obj.IgnoreCase = False
			Obj.Global = True
			Obj.Pattern = "(&?" & Name & "=)[^&]+"
			sQuery = Obj.Replace(sQuery,"$1" & Value)
			Set Obj = Nothing
		End If
	Else
		sQuery = sQuery & Name & "=" & Value
	End If
	SetQueryString = sQuery
End Function

'**********
' 函数名: GetGUID
' 作用: 生成GUID
'**********
Function GetGUID()
	On Error Resume Next
	GetGUID = Mid(CreateObject("Scriptlet.TypeLib").Guid,2,36)
	Err.Clear
	On Error Goto 0
End Function



%>
