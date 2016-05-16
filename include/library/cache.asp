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
Class Class_Cache
	Public	Mark	'前缀

	Private IExpires

    Public Default Property Get Constructor(Value)
        Constructor = [Get](Value)
    End Property

    Private Property Let Timeout(Value)
        IExpires = DateAdd("n", Value, Now)	'分钟
    End Property

    Private Property Get Timeout()
        Timeout = IExpires
    End Property

    '**********
    ' 函数名: class_Initialize
    ' 作  用: Constructor
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
    ' 函数名: lock
    ' 作  用: lock the applaction
    '**********
	Sub Lock()
        Application.Lock()
    End Sub

    '**********
    ' 函数名: UnLock
    ' 作  用: unLock the applaction
    '**********
	Sub UnLock()
        Application.unLock()
    End Sub

    '**********
    ' 函数名: SetCache
    ' 作  用: Set a cache
    '**********
	Sub [Set](Key, Value, Expire)
        Expires = Expire
        Lock
		Application(Mark & Key) = Value
        Application(Mark & Key & "_Expires") = Expires
        unLock
    End Sub

    '**********
    ' 函数名: Get
    ' 作  用: Get a cache
    '**********
	Function [Get](Key)
        Dim Expire
        Expire = Application(Mark & Key & "_Expires")
        If IsNull(Expire) Or IsEmpty(Expire) Then
            [Get] = ""
        Else
            If IsDate(Expire) And CDate(Expire) > Now Then
                [Get] = Application(Mark & Key)
            Else
                Call Remove(Mark & Key)
                Value = ""
            End If
        End If
    End Function

    '**********
    ' 函数名: remove
    ' 作  用: remove a cache
    '**********
	Sub Remove(Key)
        Lock
        Application.Contents.Remove(Mark & Key)
        Application.Contents.Remove(Mark & Key & "_Expires")
        unLock
    End Sub

    '**********
    ' 函数名: removeAll
    ' 作  用: remove all cache
    '**********
	Sub RemoveAll()
        Lock
        Application.Contents.RemoveAll()
        unLock
    End Sub

    '**********
    ' 函数名: compare
    ' 作  用: Compare two caches
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