<%
'**********
'	class		: Md5 
'	File Name	: mg5.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-5-20
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********
Class Class_Email
	Public Charset,MailServer

	Private Sub Class_Initialize()
		Charset			= "gb2312"
		MailServer			= "127.0.0.1"
	End Sub

	'**********
	'函数名：Send
	'作  用：用jMail组件发送邮件
	'参  数：MailtoEmails ----	收信人(xxx@gmail.com|爸爸    or      Array("xxx@gmail.com|爸爸","yyy@gmail.com|妈妈"))
	'        Subject       -----主题
	'        TemplateFile  -----模板文件路径
	'        Params		   -----替换参数
	'        FromName      -----发信人姓名
	'        MailFrom      -----发信人地址
	'        UserName      -----发信人帐号
	'        Password      -----发信人密码
	'        Priority      -----信件优先级（1为加急，3为普通，5为低级）
	'返回值：Array(是否成功,提示信息)
	'**********
	Public Function Send(MailtoEmails, Subject, TemplateFile, Params, FromName, FromMail, UserName, Password, Priority)
		If IsObjInstalled("jMail.Message") Then
			On Error Resume Next
		Else
			Send = Array(False,"未找到 JMail 组件")
			Exit Function
		End If
		If Priority = "" Or IsNull(Priority) Then Priority = 3

		Dim jMail : Set jMail = Server.CreateObject("jMail.Message")
		jMail.Silent = False
		jMail.Charset = Me.Charset							'邮件编码
		jMail.ContentType = "text/html"					'邮件正文格式
		jMail.From = FromMail							'发信人Email
		jMail.FromName = FromName						'发信人姓名
		jMail.ReplyTo = FromMail						
		jMail.AppendBodyFromFile TemplateFile			'内容模板
		Dim iParamName
		If Not IsNull(Params) Then
			For Each iParamName In Params
				jMail.Body = Replace(jMail.Body, "{" & iParamName & "}", Params(iParamName))
			Next
		End If

		jMail.ClearRecipients
		Dim aMailtoEmail, i
		If IsArray(MailtoEmails) Then
			For i = 0 to UBound(MailtoEmails)
				aMailtoEmail = Split(MailtoEmails(i),"|")
				If UBound(aMailtoEmail) > 0 Then
					jMail.AddRecipient aMailtoEmail(0),aMailtoEmail(1)						'收信人
				Else
					jMail.AddRecipient aMailtoEmail(0),Left(aMailtoEmail(0),InStr(aMailtoEmail(0),"@")-1)
				End If
			Next
		Else
			aMailtoEmail = Split(MailtoEmails,"|")
			If UBound(aMailtoEmail) > 0 Then
				jMail.AddRecipient aMailtoEmail(0),aMailtoEmail(1)						'收信人
			Else
				jMail.AddRecipient aMailtoEmail(0),Left(aMailtoEmail(0),InStr(aMailtoEmail(0),"@")-1)
			End If
		End If
		jMail.Subject = Subject							'主题
		jMail.Priority = Priority						'邮件等级，1为加急，3为普通，5为低级

		'如果服务器需要SMTP身份验证则还需指定以下参数
		jMail.MailDomain =	Right(FromMail,InStr(FromMail,"@"))				'域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
		jMail.MailServerUserName = UserName			'登录用户名
		jMail.MailServerPassWord = Password			'登录密码

		jMail.Send Me.MailServer
		Send = jMail.ErrorMessage
		jMail.Close
		Set jMail = Nothing
		If Err Then
			Send = Array(False,Err.Description)
			Err.Clear
			Exit Function
		End If
		Send = Array(True,Send)
		On Error Goto 0
	End Function

	'**********
	'函数名：IsObjInstalled
	'作  用：检查组件是否已经安装
	'参  数：strClassString ----组件名
	'返回值：True  ----已经安装
	'       False ----没有安装
	'**********
	Private Function IsObjInstalled(strClassString)
		IsObjInstalled = False
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If Err.Number = 0 Then IsObjInstalled = True
		Set xTestObj = Nothing
		On Error Goto 0
	End Function
End Class

%>
