<%
'**********
'	class		: Validate Code
'	File Name	: getcode.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-6-27
'**********


'**********
'	ʾ��
'**********

'********** 

'**********
'	������
'**********

Class Class_ValidCode
	'**********
	'��������GetCode
	'��  �ã���ȡ��֤�������ؼ�
	'����:lang	-- en|cn|int
	'**********
	Sub GetCode(sPath,lang)
		Dim tmpstr
		Randomize
		tmpstr=cstr(Int(900000*rnd)+100000)
		Response.Write "<img id=""codeimg"" src="""&sPath&"include/helper/validcode_"&lang&".asp?s=" & tmpstr & """ style=""cursor:pointer;border:1px solid #ccc;vertical-align:middle;"" onclick=""this.src=this.src+'&t='+ Math.random()"" alt=""&#30475;&#19981;&#28165;? &#28857;&#20987;&#21047;&#26032;"" /><input type=""hidden"" name=""codename"" id=""codename"" value=""" & tmpstr & """ />"
	End Sub
	
	'**********
	'��������CodePass
	'��  �ã������֤���Ƿ���ȷ
	'**********
	Function Check(ByVal CodeStr)
		Dim codename
		codename = Trim(Request("codename"))
		If CStr(Session("GetCode" & codename)) = CStr(CodeStr) And CodeStr <> "" Then
			Check = True
		Else
			Check = False
		End If
		Session.Contents.Remove("GetCode" & codename)
	End Function
End Class
%>
