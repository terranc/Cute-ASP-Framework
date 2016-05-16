<%
'**********
'	class		: Microsoft Debugging class
'	File Name	: Debug.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-5-17
'**********


'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Debug

    Private blnEnabled
    Private dteRequestTime
    Private dteFinishTime
    Private objStorage

    '**********
    ' 方法名: Open
    ' 作  用: 打开Debug
    '**********
    Sub Open()
        blnEnabled = true
    End Sub

    '**********
    ' 方法名: Close
    ' 作  用: 停止Debug
    '**********
    Sub Close()
        blnEnabled = false
    End Sub

    Private Sub Class_Initialize()
        dteRequestTime = Now()
        Set objStorage = Server.CreateObject("Scripting.Dictionary")
    End Sub

    '**********
    ' 方法名: Add
    ' 作  用: 添加监听变量
    '**********
    Sub Add(label)
        If blnEnabled Then
            objStorage.Add ValidLabel(label), eval(label)
        End If
    End Sub


    '**********
    ' 方法名: Show
    ' 作  用: 显示结果
    '**********
    Sub Show()
        dteFinishTime = Now()
        If blnEnabled Then
            PrintSummaryInfo()
            PrintCollection "VARIABLE STORAGE", objStorage
            PrintCollection "QUERYSTRING COLLECTION", Request.QueryString()
            PrintCollection "FORM COLLECTION", Request.Form()
            PrintCollection "COOKIES COLLECTION", Request.Cookies()
            PrintCollection "SESSION CONTENTS COLLECTION", Session.Contents()
            PrintCollection "SERVER VARIABLES COLLECTION", Request.ServerVariables()
            PrintCollection "APPLICATION CONTENTS COLLECTION", Application.Contents()
            PrintCollection "APPLICATION STATICOBJECTS COLLECTION", Application.StaticObjects()
            PrintCollection "SESSION STATICOBJECTS COLLECTION", Session.StaticObjects()
        End If
    End Sub

	Private Function ValidLabel(byval label)
        Dim i, lbl
        i = 0
        lbl = label
        Do
            If Not objStorage.Exists(lbl) Then Exit Do
            i = i + 1
            lbl = label & "(" & i & ")"
        Loop Until i = i
        ValidLabel = lbl
    End Function

    Private Sub PrintSummaryInfo()
        With Response
			.Write("<div style=""font-size:12px;"">")
            .Write("<hr size=""1"" />")
            .Write("<h3 style=""background:#7EA5D7;padding:4px;color:white;font-weight:300;font-size:12px;"">SUMMARY INFO</h3>")
            .Write("<div style=""font-size:12px;"">Time of Request = " & dteRequestTime) & "</div>"
            .Write("<div style=""font-size:12px;"">Time Finished = " & dteFinishTime) & "</div>"
            .Write("<div style=""font-size:12px;"">Elapsed Time = " & DateDiff("s", dteRequestTime, dteFinishTime) & " seconds</div>")
            .Write("<div style=""font-size:12px;"">Request Type = " & Request.ServerVariables("REQUEST_METHOD") & "</div>")
            .Write("<div style=""font-size:12px;"">Status Code = " & Response.Status & "</div>")
			.Write("</div>")
        End With
    End Sub

    Private Sub PrintCollection(Byval Name, Byval Collection)
        Dim varItem
        Response.Write("<h5 style=""margin:10px 0 0;padding:4px;background:#7EA5D7;color:white;font-weight:300;"">" & Name & "</h5>")
        For Each varItem in Collection
            Response.Write("<p style=""margin:5px 0 0;font-size:12px;"">" & varItem & "=" & Collection(varItem) & "</p>")
        Next
    End Sub

    Private Sub Class_Terminate()
        Set objStorage = Nothing
    End Sub

End Class

%>
