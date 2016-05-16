<%
'**********
'	class		: A Date class
'	File Name	: Date.asp
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

Class Class_Date
    Public TimeZone

    '**********
    ' 函数名: class_Initialize
    ' 作  用: Save the session
    '**********
    Private Sub class_initialize()
        TimeZone = 8
    End Sub

    '**********
    ' 函数名: class_Terminate
    ' 作  用: Deconstrutor
    '**********
    Private Sub class_Terminate()
    End Sub

    Private Function getMistiming(sDate)
        getMistiming = DateDiff("s", "1970-1-1 00:00:00", DateAdd("h", Me.TimeZone, sDate))
    End Function

    '**********
    ' 函数名: ToGMTdate
    ' 参  数: sDate
    ' 作  用: 本地时间转GMT时间
    '**********
    Function ToGMTdate(sDate)
        Dim dWeek, dMonth
        Dim strZero, strZone
        strZero = "00"
        If Me.TimeZone > 0 Then
            strZone = "+"&Right("0"&Me.TimeZone, 2)&"00"
        Else
            strZone = "-"&Right("0"&Me.TimeZone, 2)&"00"
        End If
        dWeek = Array("Sun", "Mon", "Tue", "Wes", "Thu", "Fri", "Sat")
        dMonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        ToGMTdate = dWeek(Weekday(sDate) -1)&", "&Right(strZero&Day(sDate), 2)&" "&dMonth(Month(sDate) -1)&" "&Year(sDate)&" "&Right(strZero&Hour(sDate), 2)&":"&Right(strZero&Minute(sDate), 2)&":"&Right(strZero&Second(sDate), 2)&" "&strZone
    End Function

    '**********
    ' 函数名: ToUnixEpoch
    ' 参  数: sDate
    ' 作  用: 获取时间戳
    '**********
    Function ToUnixEpoch(sDate)
        ToUnixEpoch = DateDiff("s", "1970-1-1 00:00:00", sDate) - getMistiming("1970-1-1 00:00:00")
    End Function

    '**********
    ' 函数名: FromUnixEpoch
    ' 参  数: iNumber	--  时间戳
    ' 作  用: 获取当地时间
    '**********
    Function FromUnixEpoch(iNumber)
        FromUnixEpoch = DateAdd("s", iNumber + getMistiming("1970-1-1 00:00:00"), "1970-1-1 00:00:00")
    End Function

    '**********
    ' 函数名: Format
    ' 参  数: sDate		--  时间
    ' 参  数: format	--  格式化格式
    ' 作  用: 格式化时间
    '**********
   Function Format(sDate, formatstr)
		Dim str
		If Len(sDate) = 0 Then Exit Function
        If Len(formatstr)>0 Then
			str = Replace(formatstr, "yyyy", Year(sDate))
			str = Replace(str, "yy", Right(Year(sDate), 2))
			str = Replace(str, "MM", doublee(Month(sDate)))
			str = Replace(str, "dd", doublee(Day(sDate)))
			str = Replace(str, "hh", doublee(Hour(sDate)))
			str = Replace(str, "mm", doublee(Minute(sDate)))
			str = Replace(str, "ss", doublee(Second(sDate)))
			str = Replace(str, "M", Month(sDate))
			str = Replace(str, "d", Day(sDate))
			str = Replace(str, "h", Hour(sDate))
			str = Replace(str, "m", Minute(sDate))
			str = Replace(str, "s", Second(sDate))
            Format = str
		Else
			Format = sDate
        End If
    End Function

    Private Function doublee(sDate)
        If Len(sDate) = 1 Then
            doublee = "0"&sDate
        Else
            doublee = sDate
        End If
    End Function

    '**********
    ' 函数名: Zodiac
    ' 参  数: bYear as birthday year
    ' 作  用: 计算所属生肖
    '**********
    Function Zodiac(bYear)
        If bYear > 0 Then
            Dim ZodiacList
            ZodiacList = Array("猴", "鸡", "狗", "猪", "鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊")
            Zodiac = ZodiacList(bYear Mod 12)
        End If
    End Function

    '**********
    ' 函数名: Constellation
    ' 参  数: Birth as birthday
    ' 作  用: 计算所属生肖
    '**********
    Function Constellation(Birth)
        If Year(Birth) <1951 Or Year(Birth) > 2049 Then Exit Function
        Dim BirthDay, BirthMonth
        BirthDay = Day(Birth)
        BirthMonth = Month(Birth)
		Dim tmp : tmp = ""
        Select Case BirthMonth
            Case 1
                If BirthDay>= 21 Then
                    tmp = tmp & "水瓶"
                Else
                    tmp = tmp & "魔羯"
                End If
            Case 2
                If BirthDay>= 20 Then
                    tmp = tmp & "双鱼"
                Else
                    tmp = tmp & "水瓶"
                End If
            Case 3
                If BirthDay>= 21 Then
                    tmp = tmp & "白羊"
                Else
                    tmp = tmp & "双鱼"
                End If
            Case 4
                If BirthDay>= 21 Then
                    tmp = tmp & "金牛"
                Else
                    tmp = tmp & "白羊"
                End If
            Case 5
                If BirthDay>= 22 Then
                    tmp = tmp & "双子"
                Else
                    tmp = tmp & "金牛"
                End If
            Case 6
                If BirthDay>= 22 Then
                    tmp = tmp & "巨蟹"
                Else
                    tmp = tmp & "双子"
                End If
            Case 7
                If BirthDay>= 23 Then
                    tmp = tmp & "狮子"
                Else
                    tmp = tmp & "巨蟹"
                End If
            Case 8
                If BirthDay>= 24 Then
                    tmp = tmp & "处女"
                Else
                    tmp = tmp & "狮子"
                End If
            Case 9
                If BirthDay>= 24 Then
                    tmp = tmp & "天秤"
                Else
                    tmp = tmp & "处女"
                End If
            Case 10
                If BirthDay>= 24 Then
                    tmp = tmp & "天蝎"
                Else
                    tmp = tmp & "天秤"
                End If
            Case 11
                If BirthDay>= 23 Then
                    tmp = tmp & "射手"
                Else
                    tmp = tmp & "天蝎"
                End If
            Case 12
                If BirthDay>= 22 Then
                    tmp = tmp & "魔羯"
                Else
                    tmp = tmp & "射手"
                End If
            Case Else
                tmp = ""
        End Select
		Constellation = tmp
    End Function

End Class
%>