<%
'''''''Ajax接口'''''''
'验证邮箱
Sub ValidateEmail(ByVal params)
	Dim xEmail
	xEmail = Tpub.rq(0,params("email"),1,"")
	If Tpub.String.Validate(xEmail,"email") Then
		die "输入正确"
	Else
		die "邮箱格式错误"
	End If
End Sub
%>