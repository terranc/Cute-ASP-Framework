<%
'**********
'	class		: A Caching class
'	File Name	: Cache.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-4-2
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********
Class Class_Cookie
	Public	Mark	'前缀

    Public Default Property Get Constructor(Value)
        Constructor = [Get](Value)
    End Property

    '**********
    ' 函数名: class_Initialize
    ' 作  用: Save the session
    '**********
	Private Sub class_initialize()
		Mark = "cute_"
    End Sub

    '**********
    ' 函数名: class_Terminate
    ' 作  用: Deconstrutor
    '**********
	Private Sub class_Terminate()
    End Sub

    '**********
    ' 函数名: Set
    ' 作  用: Add a cookie
    '**********
	Sub [Set](Key, Value, Options)
        Response.Cookies(Me.Mark & Key) = Value
        If Not (IsNull(Options) Or IsEmpty(Options)) Then
            If IsArray(Options) Then
                Dim l : l = UBound(Options)
				If l > 0 Then Response.Cookies(Me.Mark & Key).Expires = Options(0)
                If l > 1 Then Response.Cookies(Me.Mark & Key).Path = Options(1)
                If l = 2 Then Response.Cookies(Me.Mark & Key).Domain = Options(2)
            Else
                If Options <> "" Then Response.Cookies(Me.Mark & Key).Expires = Options
            End If
        End If
    End Sub

    '**********
    ' 函数名: Get
    ' 作  用: Get a cookies
    '**********
	Function [Get](Key)
        [Get] = Request.Cookies(Me.Mark & Key)
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: Remove a cookie
    '**********
	Sub Remove(Key)
         Response.Cookies(Me.Mark & Key) = Empty
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: Remove all cookies
    '**********
	Sub RemoveAll()
        Clear()
    End Sub

    '**********
    ' 函数名: Clear
    ' 作  用: Remove all cookies
    '**********
	Private Sub Clear()
        Dim iCookie
        For Each iCookie In Request.Cookies
            Response.Cookies(iCookie).Expires = formatDateTime(Now)
        Next
    End Sub

    '**********
    ' 函数名: compare
    ' 作  用: Compare two cookie
    '**********
	Function Compare(Key1, Key2)
        Dim Cache1
        Cache1 = Me.[Get](Key1)
        Dim Cache2
        Cache2 = Me.[Get](Key2)
        If TypeName(Cache1) <> TypeName(Cache2) Then
            Compare = False
        Else
            If TypeName(Cache1) = "Object" Then
                Compare = (Cache1 Is Cache2)
            Else
                If TypeName(Cache1) = "Variant()" Then
                    Compare = (Join(Cache1, "^") = Join(Cache2, "^"))
                Else
                    Compare = (Cache1 = Cache2)
                End If
            End If
        End If
    End Function

End Class
%>
