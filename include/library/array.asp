<%
'**********
'	class		: Array extensive class
'	File Name	: Array.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-4-20
'**********


'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Array

	'**********
	' 函数名: Max
	' 参  数: arr as a Array
	' 作  用: Max — 取出最大值
	'**********
	Function Max(ByVal arr)
		Dim tmp
		tmp = Me.rSort(arr)
		Max = tmp(0)
	End Function
	
	'**********
	' 函数名: Min
	' 参  数: arr as a Array
	' 作  用: Min — 取出最小值
	'**********
	Function Min(ByVal arr)
		Dim tmp
		tmp = Me.Sort(arr)
		Min = tmp(0)
	End Function
	
	'**********
	' 函数名: UnShift
	' 参  数: arr as an Array
	' 作  用: UnShift — 从前压入元素
	'**********
	Function UnShift(ByVal arr, ByVal var)
		Dim i, tmp
		tmp = Me.ToString(arr)
		tmp = var & "," & tmp
		tmp = Me.ToArray(tmp)
		UnShift = tmp
	End Function
	
	'**********
	' 函数名: Shift
	' 参  数: arr as an Array
	' 作  用: Shift — 从前删除元素
	'**********
	Function Shift(ByVal arr)
		Dim i, tmp
		tmp = ""
		For i = 0 To UBound(arr)
			If i<>0 Then tmp = tmp & arr(i) & ","
		Next
		tmp = Me.Strip(tmp)
		Shift = tmp
	End Function
	
	'**********
	' 函数名: Push
	' 参  数: arr as an Array
	' 参  数: var as a variable added to an array
	' 作  用: Push — 从后压入元素
	'**********
	Function Push(ByVal arr, ByVal var)
		Dim tmp : tmp = Me.ToString(arr)
		tmp = tmp & "," & Me.ConvComma(var)
		tmp = Me.ToArray(tmp)
		Push = tmp
	End Function
	
	'**********
	' 函数名: Pop
	' 参  数: arr as an array
	' 作  用: Pop — 从后删除元素
	'**********
	Function Pop(ByVal arr)
		Dim i, tmp
		For i = 0 To UBound(arr)
			If i<>UBound(arr) Then tmp = tmp & arr(i) & ","
		Next
		tmp = Me.Strip(tmp)
		Pop = tmp
	End Function
	
	'**********
	' 函数名: Strip
	' 参  数: str as a string such as "1,2,3,"
	' 作  用: Strip "," of string
	'**********
	Function Strip(ByVal Str)
		If IsArray(Str) Then Str = Me.ToString(Str)
		If Left(Str, 1) = "," Then Str = Right(Str, Len(Str) -1)
		If Right(Str, 1) = "," Then Str = Left(Str, Len(Str) -1)
		Str = Me.ToArray(Str)
		Strip = Str
	End Function
	
	'**********
	' 函数名: Walk
	' 参  数: arr as an Array
	' 参  数: callback as callback function
	' 作  用: Walk — 对数组内元素执行函数后返回新数组
	'**********
	Function Walk(ByVal arr, ByVal callback)
		Dim e : e = ""
		Dim tmp : tmp = ""
		For Each e in arr
			If IsArray(e) Then
				Execute("tmp=tmp&" & callback & "(""" & Me.ToString(e) & """)" & "&"",""")
			Else
				Execute("tmp=tmp&" & callback & "(""" & e & """)" & "&"",""")
			End If
		Next
		tmp = Me.Strip(tmp)
		Walk = tmp
	End Function
	
	'**********
	' 函数名: Splice
	' 参  数: arr as an array
	' 参  数: start as start index
	' 参  数: final as end index
	' 作  用: Splice — 从一个数组中移除一个或多个元素
	'**********
	Function Splice(ByVal arr, ByVal start, ByVal final)
		Dim i, temp, tmp
		If start > final Then
			temp = start
			start = final
			final = temp
		End If
		For i = 0 To UBound(arr)
			If i < start Or i > final Then tmp = tmp & arr(i) & ","
		Next
		tmp = Me.Strip(tmp)
		Splice = tmp
	End Function
	
	'**********
	' 函数名: Fill
	' 参  数: arr as a Array
	' 参  数: index as index to insert into an array
	' 参  数: value as element to insert into an array
	' 作  用: Fill — 插入元素
	'**********
	Function Fill(ByVal arr, ByVal index, ByVal Value)
		Dim i, tmp
		For i = 0 To UBound(arr)
			If i <> index Then
				tmp = tmp & arr(i) & ","
			Else
				tmp = tmp & Value & "," & arr(i) & ","
			End If
		Next
		tmp = Me.Strip(tmp)
		Fill = tmp
	End Function
	
	'**********
	' 函数名: Unique
	' 参  数: arr as a Array
	' 作  用: Unique — 移除重复的元素
	'**********
	Function Unique(ByVal arr)
		Dim tmp, e
		For Each e in arr
			If InStr(1, tmp, e) = 0 Then
				tmp = tmp & e & ","
			End If
		Next
		tmp = Me.Strip(tmp)
		Unique = tmp
	End Function

	'**********
	' 函数名: Reverse
	' 参  数: arr as a Array
	' 作  用: Reverse — 反向
	'**********
	Function Reverse(ByVal arr)
		Dim tmp, e
		For Each e in arr
			tmp = tmp & e & ","
		Next
		tmp = StrReverse(tmp)
		tmp = Me.Strip(tmp)
		Reverse = tmp
	End Function
	
	'**********
	' 函数名: Search
	' 参  数: arr as a Array
	' 参  数: value as Searching value
	' 作  用: Search — 查询元素，不存在则返回False
	'**********
	Function Search(ByVal arr, ByVal Value)
		Dim i
		For i = 0 To UBound(arr)
			If arr(i) = Value Then
				Search = i
				Exit Function
			End If
		Next
		Search = -1
	End Function
	
	'**********
	' 函数名: Rand
	' 参  数: arr as a Array
	' 参  数: num as specifies how many entries you want to pick
	' 作  用: Rand — 乱序
	'**********
	Function Rand(ByVal arr, ByVal num)
		Dim tmpi, tmp, i
		For i = 0 To num -1
			Randomize
			tmpi = Int((UBound(arr) + 1) * Rnd)
			tmp = tmp & arr(tmpi) & ","
		Next
		tmp = Me.Strip(tmp)
		Rand = tmp
	End Function
	
	'**********
	' 函数名: Sort
	' 参  数: arr as a Array
	' 作  用: Sort — 顺序
	'**********
	Function Sort(ByVal arr)
		Dim tmp, i, j
		ReDim tmpA(UBound(arr))
		For i = 0 To UBound(tmpA)
			tmpA(i) = CDbl(arr(i))
		Next
		For i = 0 To UBound(tmpA)
			For j = i + 1 To UBound(tmpA)
				If tmpA(i) > tmpA(j) Then
					tmp = tmpA(i)
					tmpA(i) = tmpA(j)
					tmpA(j) = tmp
				End If
			Next
		Next
		Sort = tmpA
	End Function
	
	'**********
	' 函数名: rSort
	' 参  数: arr as a Array
	' 作  用: rSort — 倒序
	'**********
	Function rSort(ByVal arr)
		Dim tmp, i, j
		ReDim tmpA(UBound(arr))
		For i = 0 To UBound(tmpA)
			tmpA(i) = CDbl(arr(i))
		Next
		For i = 0 To UBound(tmpA)
			For j = i + 1 To UBound(tmpA)
				If tmpA(i) < tmpA(j) Then
					tmp = tmpA(i)
					tmpA(i) = tmpA(j)
					tmpA(j) = tmp
				End If
			Next
		Next
		rSort = tmpA
	End Function

	'**********
	' 函数名: Shuffle
	' 参  数: arr as a Array
	' 作  用: Shuffle — 随机排序
	'**********
	Function Shuffle(ByVal arr)
		Dim m, n, i
		'i = Search(arr,Rand(arr,1))
		'arr = Splice(arr,i,i+1)
		Randomize   
		For i = 0 to UBound(arr)
			m = Int(Rnd()*i)
			n = arr(m)
			arr(m) = arr(i) 
			arr(i) = n
		Next
		Shuffle = arr
	End Function

	'**********
	' 函数名: ConvComma
	' 参  数: star as a string
	'**********
	Function ConvComma(ByVal str)
		ConvComma = Replace(str,",","&#44;")
	End Function

	'**********
	' 函数名: implode
	' 参  数: glue as a split character
	' 参  数: arr as a output array
	' 作  用: Join array elements with a string
	'**********
	Function ToString(ByVal arr)
		If IsArray(arr) Then
			Dim tmp
			tmp = Join(arr, ",")
			ToString = tmp
		Else
			ToString = arr
		End If
	End Function

	'**********
	' 函数名: ToArray
	' 参  数: str as a string converted to an array
	' 作  用: Convert to an array
	' Remarks: dim a : a = "a, b, c"
	'		   prinr(ToArray(a))
	'**********
	Function ToArray(ByVal str)
		Dim tmp
		tmp = Split(str, ",")
		ToArray = tmp
	End Function

End Class

%>