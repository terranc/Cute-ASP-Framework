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
Class Class_Session
	Public	Mark	'前缀

    Public Property Let Timeout(Value)
		If IsNumeric(Value) Then Session.Timeout = Value
    End Property

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
    ' 作  用: Add a Session
    '**********
	Sub [Set](Key, Value)
		If IsObject(Value) Then
			Set Session(Me.Mark & Key) = Value
		Else
			Session(Me.Mark & Key) = Value
		End If
    End Sub

    '**********
    ' 函数名: Get
    ' 作  用: Get a Session
    '**********
	Function [Get](Key)
		If IsObject(Session(Me.Mark & Key)) Then
			Set [Get] = Session(Me.Mark & Key)
		Else
			[Get] = Session(Me.Mark & Key)
		End If
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: Remove a Session
    '**********
	Sub Remove(Key)
		If IsObject(Session(Me.Mark & Key)) Then
			Set Session(Me.Mark & Key) = Nothing
		End If
        Session.Contents.Remove(Me.Mark & Key)
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: Remove all Session
    '**********
	Sub RemoveAll()
        Dim iSession
        For Each iSession In Session.Contents
			Me.Remove(Replace(iSession,Me.Mark,""))
        Next
	End Sub

    '**********
    ' 函数名: compare
    ' 作  用: Compare two session
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
