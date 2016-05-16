<%
'**********
'	class		: A Date class
'	File Name	: Date.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-4-2
'**********


'**********
'	ʾ��
'**********

'**********

'**********
'	������
'**********

Class Class_Date
    Public TimeZone

    '**********
    ' ������: class_Initialize
    ' ��  ��: Save the session
    '**********
    Private Sub class_initialize()
        TimeZone = 8
    End Sub

    '**********
    ' ������: class_Terminate
    ' ��  ��: Deconstrutor
    '**********
    Private Sub class_Terminate()
    End Sub

    Private Function getMistiming(sDate)
        getMistiming = DateDiff("s", "1970-1-1 00:00:00", DateAdd("h", Me.TimeZone, sDate))
    End Function

    '**********
    ' ������: ToGMTdate
    ' ��  ��: sDate
    ' ��  ��: ����ʱ��תGMTʱ��
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
    ' ������: ToUnixEpoch
    ' ��  ��: sDate
    ' ��  ��: ��ȡʱ���
    '**********
    Function ToUnixEpoch(sDate)
        ToUnixEpoch = DateDiff("s", "1970-1-1 00:00:00", sDate) - getMistiming("1970-1-1 00:00:00")
    End Function

    '**********
    ' ������: FromUnixEpoch
    ' ��  ��: iNumber	--  ʱ���
    ' ��  ��: ��ȡ����ʱ��
    '**********
    Function FromUnixEpoch(iNumber)
        FromUnixEpoch = DateAdd("s", iNumber + getMistiming("1970-1-1 00:00:00"), "1970-1-1 00:00:00")
    End Function

    '**********
    ' ������: Format
    ' ��  ��: sDate		--  ʱ��
    ' ��  ��: format	--  ��ʽ����ʽ
    ' ��  ��: ��ʽ��ʱ��
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
    ' ������: Zodiac
    ' ��  ��: bYear as birthday year
    ' ��  ��: ����������Ф
    '**********
    Function Zodiac(bYear)
        If bYear > 0 Then
            Dim ZodiacList
            ZodiacList = Array("��", "��", "��", "��", "��", "ţ", "��", "��", "��", "��", "��", "��")
            Zodiac = ZodiacList(bYear Mod 12)
        End If
    End Function

    '**********
    ' ������: Constellation
    ' ��  ��: Birth as birthday
    ' ��  ��: ����������Ф
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
                    tmp = tmp & "ˮƿ"
                Else
                    tmp = tmp & "ħ��"
                End If
            Case 2
                If BirthDay>= 20 Then
                    tmp = tmp & "˫��"
                Else
                    tmp = tmp & "ˮƿ"
                End If
            Case 3
                If BirthDay>= 21 Then
                    tmp = tmp & "����"
                Else
                    tmp = tmp & "˫��"
                End If
            Case 4
                If BirthDay>= 21 Then
                    tmp = tmp & "��ţ"
                Else
                    tmp = tmp & "����"
                End If
            Case 5
                If BirthDay>= 22 Then
                    tmp = tmp & "˫��"
                Else
                    tmp = tmp & "��ţ"
                End If
            Case 6
                If BirthDay>= 22 Then
                    tmp = tmp & "��з"
                Else
                    tmp = tmp & "˫��"
                End If
            Case 7
                If BirthDay>= 23 Then
                    tmp = tmp & "ʨ��"
                Else
                    tmp = tmp & "��з"
                End If
            Case 8
                If BirthDay>= 24 Then
                    tmp = tmp & "��Ů"
                Else
                    tmp = tmp & "ʨ��"
                End If
            Case 9
                If BirthDay>= 24 Then
                    tmp = tmp & "���"
                Else
                    tmp = tmp & "��Ů"
                End If
            Case 10
                If BirthDay>= 24 Then
                    tmp = tmp & "��Ы"
                Else
                    tmp = tmp & "���"
                End If
            Case 11
                If BirthDay>= 23 Then
                    tmp = tmp & "����"
                Else
                    tmp = tmp & "��Ы"
                End If
            Case 12
                If BirthDay>= 22 Then
                    tmp = tmp & "ħ��"
                Else
                    tmp = tmp & "����"
                End If
            Case Else
                tmp = ""
        End Select
		Constellation = tmp
    End Function

End Class
%>