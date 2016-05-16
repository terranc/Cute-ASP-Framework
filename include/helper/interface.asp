<%
'**********
'	class		: Interface
'	File Name	: Class_Interface.asp
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
Class Class_Interface
	Private s_stm
		
	'����Stream�������
	Public Property Let Stream(str)
		s_stm = str
	End Property

	'**********
    ' ������: class_Initialize
    ' ��  ��: Constructor
    '**********
	Private Sub Class_Initialize()
		s_stm = "ADODB.Stream"
    End Sub
	
    '**********
    '��������PostHttpPage
	'��  ����RefererUrl	---- ����ҳ��
	'		 PostUrl	---- ��ȡ��ַ
	'		 PostData	---- ���Ͳ���
	'		 DataType	---- ��������
    '��  �ã���¼
    '**********
    Function PostHttpPage(RefererUrl, PostUrl, PostData, DataType)
        Dim xmlHttp, RetStr
		If PostUrl = "" Then Exit Function
        On Error Resume Next
        Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		With XmlHttp
			.SetTimeouts 10000, 10000, 10000, 10000
			.Open "POST", PostUrl, false
			.setRequestHeader "Content-Length", Len(PostData)
			.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			.setRequestHeader "Referer", RefererUrl
			.Send PostData
		End With
        If Err.Number <> 0 Then
            Set xmlHttp = Nothing
            PostHttpPage = "$False$"
            Exit Function
        End If
		On Error Goto 0
        PostHttpPage = bytesToBSTR(xmlHttp.responseBody, DataType)
        Set xmlHttp = Nothing
    End Function

    '**********
    '��������SaveRemoteFile
    '��  �ã�����Զ�̵��ļ�������
    '��  ����LocalFileName ------ �����ļ���
    '��  ����RemoteFileUrl ------ Զ���ļ�URL
    '��  ����Referer ------ Զ�̵����ļ����Ը����ɼ��ģ�������ҳ��ַ��û�з������գ�
    '**********
    Function SaveRemoteFile(LocalFileName, RemoteFileUrl, Referer)
		SaveRemoteFile = False
		If RemoteFileUrl = "" Then Exit Function
		SaveRemoteFile = True
        Dim Ads, Retrieval, GetRemoteData
        On Error Resume Next
        Set Retrieval = Server.CreateObject("Msxml2.XMLHTTP")
        With Retrieval
			'.SetTimeouts 10000, 10000, 10000, 10000
            .Open "Get", RemoteFileUrl, False, "", ""
            .Send
            If .Readystate<>4 Or .Status > 300 Then
                SaveRemoteFile = False
                Exit Function
            End If
            GetRemoteData = .ResponseBody
        End With
        Set Retrieval = Nothing
        Set Ads = Server.CreateObject(s_stm)
        With Ads
            .Type = 1
            .Open
            .Write GetRemoteData
            .SaveToFile Server.MapPath(LocalFileName), 2
            .Cancel()
            .Close()
        End With
        If Err.Number<>0 Then
            SaveRemoteFile = False
			On Error Goto 0
            Exit Function
        End If
        Set Ads = Nothing
    End Function

    '**********
    '��������GetHttpPage
    '��  �ã���ȡ��ҳԴ��
    '��  ����HttpUrl ------��ҳ��ַ,Cset ����
    '**********
    Function GetHttpPage(URL, Cset, iUserName , iPassword)
        Dim xmlHttp
        If URL = "" Or Len(URL)<18 Or URL = "$False$" Then
            GetHttpPage = "$False$"
            Exit Function
        End If
        Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
		With xmlHttp
			.SetTimeouts 10000, 10000, 10000, 10000
			.Open "GET", URL, False, iUserName, iPassword
			.Send()
		End With
        If xmlHttp.Readystate<>4 Then
            Set xmlHttp = Nothing
            GetHttpPage = "$False$"
            Exit Function
        End If
        GetHTTPPage = bytesToBSTR(xmlHttp.responseBody, Cset)
        Set xmlHttp = Nothing
    End Function

    '**********
    '��������BytesToBstr
    '��  �ã�����ȡ��Դ��ת��Ϊ����
    '��  ����Body ------Ҫת���ı���
    '��  ����Cset ------Ҫת��������
    '**********
    Private Function BytesToBstr(Body, Cset)
        Dim Objstream
        Set Objstream = Server.CreateObject(s_stm)
		With objstream
			.Type = 1
			.Mode = 3
			.Open
			.Write body
			.Position = 0
			.Type = 2
			.Charset = Cset
			BytesToBstr = objstream.ReadText
			.Close
		End With
        Set objstream = Nothing
    End Function
	
	'**********
	'��������GetBody
	'��  �ã���ȡ�ַ���
	'��  ����ConStr ------��Ҫ��ȡ���ַ���
	'��  ����StartStr ------��ʼ�ַ���
	'��  ����OverStr ------�����ַ���
	'��  ����IncluL ------�Ƿ����StartStr
	'��  ����IncluR ------�Ƿ����OverStr
	'**********
	Function GetBody(ConStr,StartStr,OverStr,IncluL,IncluR)
	   If ConStr = "$False$" or ConStr = "" or IsNull(ConStr) = True Or StartStr = "" or IsNull(StartStr) = True Or OverStr = "" or IsNull(OverStr) = True Then
		  GetBody = "$False$"
		  Exit Function
	   End If
	   Dim ConStrTemp
	   Dim Start,Over
	   ConStrTemp = Lcase(ConStr)
	   StartStr = Lcase(StartStr)
	   OverStr = Lcase(OverStr)
	   Start  =  InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
	   If Start <= 0 then
		  GetBody = "$False$"
		  Exit Function
	   Else
		  If IncluL = False Then
			 Start = Start+LenB(StartStr)
		  End If
	   End If
	   Over = InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
	   If Over <= 0 Or Over <= Start then
		  GetBody = "$False$"
		  Exit Function
	   Else
		  If IncluR = True Then
			 Over = Over+LenB(OverStr)
		  End If
	   End If
	   GetBody = MidB(ConStr,Start,Over-Start)
	End Function


	'**********
	'��������GetLinkArray
	'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
	'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
	'��  ����StartStr ------��ʼ�ַ���
	'��  ����OverStr ------�����ַ���
	'��  ����IncluL ------�Ƿ����StartStr
	'��  ����IncluR ------�Ƿ����OverStr
	'**********
	Function GetLinkArray(ConStr,StartStr,OverStr,IncluL,IncluR)
	   If ConStr = "$False$" or ConStr = "" Or IsNull(ConStr) = True or StartStr = "" Or OverStr = "" or  IsNull(StartStr) = True Or IsNull(OverStr) = True Then
		  GetLinkArray = "$False$"
		  Exit Function
	   End If
	   Dim TempStr,TempStr2,objRegExp,Matches,Match
	   TempStr = ""
	   Set objRegExp  =  New Regexp 
	   objRegExp.IgnoreCase  =  True 
	   objRegExp.Global  =  True
	   objRegExp.Pattern  =  "("&StartStr&").+?("&OverStr&")"
	   Set Matches  = objRegExp.Execute(ConStr) 
	   For Each Match in Matches
		  TempStr = TempStr & "$Array$" & Match.Value
	   Next 
	   Set Matches = nothing
	
	   If TempStr = "" Then
		  GetLinkArray = "$False$"
		  Exit Function
	   End If
	   TempStr = Right(TempStr,Len(TempStr)-7)
	   If IncluL = False then
		  objRegExp.Pattern  = StartStr
		  TempStr = objRegExp.Replace(TempStr,"")
	   End if
	   If IncluR = False then
		  objRegExp.Pattern  = OverStr
		  TempStr = objRegExp.Replace(TempStr,"")
	   End if
	   Set objRegExp = nothing
	   Set Matches = nothing
	   
	   TempStr = Replace(TempStr,"""","")
	   TempStr = Replace(TempStr,"'","")
	   TempStr = Replace(TempStr," ","")
	   TempStr = Replace(TempStr,"(","")
	   TempStr = Replace(TempStr,")","")
	
	   If TempStr = "" then
		  GetLinkArray = "$False$"
	   Else
		  GetLinkArray = TempStr
	   End if
	End Function
End Class
%>