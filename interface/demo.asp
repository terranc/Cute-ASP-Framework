<%
'''''''Ajax�ӿ�'''''''
'��֤����
Sub ValidateEmail(ByVal params)
	Dim xEmail
	xEmail = Tpub.rq(0,params("email"),1,"")
	If Tpub.String.Validate(xEmail,"email") Then
		die "������ȷ"
	Else
		die "�����ʽ����"
	End If
End Sub
%>