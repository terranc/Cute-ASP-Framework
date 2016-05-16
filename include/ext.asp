<%
'**********
'	class		: A Extensive function Library
'	File Name	: ext.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-4-3
'**********
'**********
'	ʾ��
'**********

'**********
' ������: echo
' ��  ��: str as a output string
' ��  ��: Print the value of a variable
'**********
Sub echo(ByVal str)
    Response.Write str
End Sub

'********** 
' ������: die
' Param: str as a output string
' ����: Print the value of a variable and exit the procedure
'********** 
Sub die(str)
	echo(str)
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	Response.End()
End Sub

'**********
' ������: isset
' ��  ��: Obj as a object
' ��  ��: isNothing �� Check if the object is nothing or null or undefined
'**********
Function isset(Obj)
    isset = true
    If IsEmpty(Obj) Then
        isset = false
        Exit Function
    End If
    If IsNull(Obj) Then
        isset = false
        Exit Function
    End If
	If IsObject(Obj) Then
		If Obj Is Nothing Then
			isset = false
			Exit Function
		End If
	Else
		If Not IsArray(Obj) Then
			If Obj = "" Then
				isset = false
				Exit Function
			End If
		End If
	End If
End Function

'**********
'��������isNumber
'��  �ã��ж��Ƿ�����
'**********
Function isNumber(str)
	isNumber = False
	If isset(str) Then
		If isNumeric(str) Then
			isNumber = True
		End if
	End If
End Function

'**********
'��������locationHref
'��  �ã�ҳ����ת
'**********
Sub locationHref(url)
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	Response.Redirect url
	die("")
End Sub

'**********
'��������Referer
'��  �ã�������ҳ
'**********
Sub locationReferer()
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	locationHref(Request.ServerVariables("HTTP_REFERER"))
	die("")
End Sub

'**********
'��������AlertRedirect
'��  �ã���Ϣ��
'**********
Sub alertRedirect(msgstr,url)
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	die "<script type=""text/javascript"">"&vbCr& _
			"alert("""&msgstr&""");"&vbCr& _
			"location.replace("""&url&""");"&vbCr& _
			"</script>"
End Sub

'**********
'��������AlertBack
'��  ����msgstr	-- ������Ϣ
'��  �ã���Ϣ��
'**********
Sub alertBack(msgstr)
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	die "<script type=""text/javascript"">alert(""" & msgstr & """);history.back();</script>"
End Sub

'**********
'��������AlertClose
'��  ����msgstr	-- ������Ϣ
'��  �ã���Ϣ���رմ���
'**********
Sub alertClose(msgstr)
	On Error Resume Next
	Tpub.db.closeRs rs
	Set Tpub = Nothing
	die "<script type=""text/javascript"">alert(""" & msgstr & """);window.close();</script>"
End Sub

'********** 
' ������: IIf
' ����: ����ֵ�жϽ��
'********** 
Function IIf(flag,return1,return2)
	If flag Then
		IIf = return1
	Else
		IIf = return2
	End If
End Function

'********** 
' ����: ReAopResult
' ���ã�������Ϣ�洢������
'********** 
Class ReAopResult
	Public Code
	Public Message
	Public AttachObject
End Class
%>