<%
'**********
'	class		: A Item class
'	File Name	: Item.asp
'	Version		: 0.2.0
'	Author		: TerranC
'	Date		: 2008-6-16
'**********


'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Params
 	'**********
    ' 函数名: Contents
    ' 作  用: Get Params Value
    '**********
    Public Default Property Get Constructor(OutParams)
		Set OutParams = CreateObject("Scripting.Dictionary")
		OutParams.CompareMode = 1
    End Property

	'**********
    ' 函数名: class_Initialize
    ' 作  用: Constructor
    '**********
	Private Sub Class_Initialize()
    End Sub

	'**********
    ' 函数名: class_Initialize
    ' 作  用: Constructor
    '**********
	Private Sub Class_Terminate()
    End Sub

	Sub Close(OutParams)
		Set OutParams = Nothing
	End Sub
End Class
%>