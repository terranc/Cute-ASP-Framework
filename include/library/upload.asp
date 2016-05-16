<%
'**********
'	class		: upload
'	File Name	: upload.asp
'	Version		: 0.2.0
'	Updater		: TerranC
'	Date		: 2008-11-25
'**********


'**********
'	示例
'**********

'**********
'设置上传模式
'属性Mode: 0.无组件上传	1.AspUpload上传

'**********
'	构建类
'**********

Class Class_Upload
    Private objForm, binForm, binItem, strDate, lngTime, objFormASP
    Public Error '出错信息
    Public MaxSize '单文件最大上传大小
    Public TotalSize '总文件最大上传大小
    Public Mode '上传模式
    Public SavePath '上传路径
    Public Charset '字符集
    Public FileType '允许上传的文件类型
    Public AutoSave '文件名保存方式
    Public FormItem, FileItem
    Public UpCount	'总上传文件数

    Private Sub Class_Initialize
        Error = -1
        MaxSize = 10485760 '默认为10mb
        FileType = "jpg/gif/jpeg/png/bmp"
        SavePath = "./upfile/"
        AutoSave = 0
        TotalSize = 0
        Mode = 0
		UpCount = 0
        Charset = "gb2312"
        strDate = Replace(CStr(Year(Date()) & "-" & Month(Date()) & "-" & Day(Date())), "-", "")
        lngTime = CLng(Timer() * 1000)
        Set binForm = Server.CreateObject("ADODB.Stream")
        Set binItem = Server.CreateObject("ADODB.Stream")
        Set objForm = Server.CreateObject("Scripting.Dictionary")
        objForm.CompareMode = 1
    End Sub

    Private Sub Class_Terminate
        Me.Close()
    End Sub

    Sub Close()
        objForm.RemoveAll
        Set objForm = Nothing
        Set binItem = Nothing
        If Me.Mode = 1 Then Set objFormASP = Nothing
        If Error<>4 Then binForm.Close()
        Set binForm = Nothing
    End Sub

    Sub Open()
        If Error = -1 Then
            Error = 0
        Else
            Exit Sub
        End If
        Dim lngRequestSize, binRequestData, strFormItem, strFileItem
        If Me.Mode = 1 Then
            On Error Resume Next
            Set objFormASP = Server.CreateObject("Persits.Upload")
            objFormASP.OverWriteFiles = True '同名覆盖
            objFormASP.IgnorenoPost = True
            objFormASP.SetMaxSize Me.MaxSize, True '设置单个文件最大上传
            objFormASP.CreateDirectory Server.MapPath(Me.SavePath), True '建立目录
            If LCase(Me.Charset) = "gb2312" Then
                objFormASP.CodePage = 936
            Else
                objFormASP.CodePage = 65001
            End If
            UpCount = objFormASP.Save()
            For Each strFormItem in objFormASP.Form
                objForm.Add strFormItem.Name, strFormItem.Value
            Next
            If UpCount > 0 Then
                For Each strFileItem in objFormASP.Files
                    With objForm
                        .Add strFileItem.Name, strFileItem.FileName
                        .Add strFileItem.Name&"_From", strFileItem.Binary
                        .Add strFileItem.Name&"_Name", strFileItem.OriginalFileName
                        .Add strFileItem.Name&"_Type", strFileItem.ContentType
                        .Add strFileItem.Name&"_Ext", Mid(LCase(strFileItem.Ext), 2, Len(LCase(strFileItem.Ext)))
                        .Add strFileItem.Name&"_Size", strFileItem.Size
                        .Add strFileItem.Name&"_Path", strFileItem.Path
                        .Add strFileItem.Name&"_Width", strFileItem.ImageWidth 'image only
                        .Add strFileItem.Name&"_Height", strFileItem.ImageHeight 'image only
                    End With
                    lngRequestSize = lngRequestSize + strFileItem.Size
                    If Me.AutoSave<>2 Then
                        intTemp = GetFerr(objForm(strFileItem.Name&"_Ext"), 2, Len(objForm(strFileItem.Name&"_Ext")))
                        objForm.Add strFileItem.Name&"_Err", intTemp
                        If intTemp = 0 Then
                            If Me.AutoSave = 0 Then
                                strFnam = GetTimeStr()
                                If objForm(strFileItem.Name&"_Ext")<>"" Then strFnam = strFnam&"."&objForm(strFileItem.Name&"_Ext")
                            Else
                                strFnam = objForm(strFileItem.Name&"_Name")
                            End If
                            objFormASP.Files(strFileItem.Name).SaveAS Server.MapPath(Me.SavePath&strFnam)
                            objForm(strFileItem.Name&"_Path") = objFormASP.Files(strFileItem.Name).Path
                        End If
                    End If
                Next
            End If
            If lngRequestSize>Me.TotalSize And Me.TotalSize<>0 Then
                Error = 4
                Exit Sub
            End If
            On Error GoTo 0
        Else
            Const strSplit = "'"">"
            lngRequestSize = Request.TotalBytes
            If lngRequestSize<1 Or (lngRequestSize>Me.TotalSize And Me.TotalSize<>0) Then
                Error = 4
                Exit Sub
            End If
            binRequestData = Request.BinaryRead(lngRequestSize)
            binForm.Type = 1
            binForm.Open
            binForm.Write binRequestData

            Dim bCrLf, strSeparator, intSeparator
            bCrLf = ChrB(13)&ChrB(10)

            intSeparator = InstrB(1, binRequestData, bCrLf) -1
            strSeparator = LeftB(binRequestData, intSeparator)

            Dim p_start, p_end, strItem, strInam, intTemp, strTemp
            Dim strFtyp, strFnam, strFext, lngFsiz
            p_start = intSeparator + 2
            Do
                p_end = InStrB(p_start, binRequestData, bCrLf&bCrLf) +3
                binItem.Type = 1
                binItem.Open
                binForm.Position = p_start
                binForm.CopyTo binItem, p_end - p_start
                binItem.Position = 0
                binItem.Type = 2
                binItem.Charset = Me.Charset
                strItem = binItem.ReadText
                binItem.Close()

                p_start = p_end
                p_end = InStrB(p_start, binRequestData, strSeparator) -1
                binItem.Type = 1
                binItem.Open
                binForm.Position = p_start
                lngFsiz = p_end - p_start -2
                binForm.CopyTo binItem, lngFsiz

                intTemp = InStr(39, strItem, """")
                strInam = Mid(strItem, 39, intTemp -39)
                If InStr(intTemp, strItem, "filename=""")<>0 Then
                    If Not objForm.Exists(strInam&"_From") Then
						UpCount = UpCount + 1
                        strFileItem = strFileItem&strSplit&strInam
                        If binItem.Size<>0 Then
                            intTemp = intTemp + 13
                            strFtyp = Mid(strItem, InStr(intTemp, strItem, "Content-Type: ") + 14)
                            strTemp = Mid(strItem, intTemp, InStr(intTemp, strItem, """") - intTemp)
                            intTemp = InstrRev(strTemp, "\")
                            strFnam = Mid(strTemp, intTemp + 1)
                            objForm.Add strInam&"_Type", strFtyp
                            objForm.Add strInam&"_Name", strFnam
                            objForm.Add strInam&"_Path", Left(strTemp, intTemp)
                            objForm.Add strInam&"_Size", lngFsiz
                            If InStr(strFnam, ".")<>0 Then
                                strFext = Mid(strTemp, InstrRev(strTemp, ".") + 1)
                            Else
                                strFext = ""
                            End If
                            If Left(strFtyp, 6) = "image/" Then
                                binItem.Position = 0
                                binItem.Type = 1
                                strTemp = binItem.Read(10)
                                If StrComp(strTemp, chrb(255) & chrb(216) & chrb(255) & chrb(224) & chrb(0) & chrb(16) & chrb(74) & chrb(70) & chrb(73) & chrb(70), 0) = 0 Then
                                    If LCase(strFext)<>"jpg" Then strFext = "jpg"
                                    binItem.Position = 3
                                    Do While Not binItem.EOS
                                        Do
                                            intTemp = ascb(binItem.Read(1))
                                        Loop While intTemp = 255 And Not binItem.EOS
                                        If intTemp < 192 Or intTemp > 195 Then
                                            binItem.Read(Bin2Val(binItem.Read(2)) -2)
                                        Else
                                            Exit Do
                                        End If
                                        Do
                                            intTemp = ascb(binItem.Read(1))
                                        Loop While intTemp < 255 And Not binItem.EOS
                                    Loop
                                    binItem.Read(3)
                                    objForm.Add strInam&"_Height", Bin2Val(binItem.Read(2))
                                    objForm.Add strInam&"_Width", Bin2Val(binItem.Read(2))
                                ElseIf StrComp(leftB(strTemp, 8), chrb(137) & chrb(80) & chrb(78) & chrb(71) & chrb(13) & chrb(10) & chrb(26) & chrb(10), 0) = 0 Then
                                    If LCase(strFext)<>"png" Then strFext = "png"
                                    binItem.Position = 18
                                    objForm.Add strInam&"_Width", Bin2Val(binItem.Read(2))
                                    binItem.Read(2)
                                    objForm.Add strInam&"_Height", Bin2Val(binItem.Read(2))
                                ElseIf StrComp(leftB(strTemp, 6), chrb(71) & chrb(73) & chrb(70) & chrb(56) & chrb(57) & chrb(97), 0) = 0 Or StrComp(leftB(strTemp, 6), chrb(71) & chrb(73) & chrb(70) & chrb(56) & chrb(55) & chrb(97), 0) = 0 Then
                                    If LCase(strFext)<>"gif" Then strFext = "gif"
                                    binItem.Position = 6
                                    objForm.Add strInam&"_Width", BinVal2(binItem.Read(2))
                                    objForm.Add strInam&"_Height", BinVal2(binItem.Read(2))
                                ElseIf StrComp(leftB(strTemp, 2), chrb(66) & chrb(77), 0) = 0 Then
                                    If LCase(strFext)<>"bmp" Then strFext = "bmp"
                                    binItem.Position = 18
                                    objForm.Add strInam&"_Width", BinVal2(binItem.Read(4))
                                    objForm.Add strInam&"_Height", BinVal2(binItem.Read(4))
                                End If
							End If
							objForm.Add strInam&"_Ext", strFext
							objForm.Add strInam&"_From", p_start
							intTemp = GetFerr(lngFsiz, strFext)
							If Me.AutoSave<>2 Then
								objForm.Add strInam&"_Err", intTemp
								If intTemp = 0 Then
									If Me.AutoSave = 0 Then
										strFnam = GetTimeStr()
										If strFext<>"" Then strFnam = strFnam&"."&strFext
									End If
									binItem.SaveToFile Server.MapPath(Me.SavePath&strFnam), 2
									objForm.Add strInam, strFnam
								End If
							End If
						Else
							UpCount = 0
							objForm.Add strInam&"_Err", -1
						End If
					End If
				Else
					binItem.Position = 0
					binItem.Type = 2
					binItem.Charset = Me.Charset
					strTemp = binItem.ReadText
					If objForm.Exists(strInam) Then
						objForm(strInam) = objForm(strInam)&","&strTemp
					Else
						strFormItem = strFormItem&strSplit&strInam
						objForm.Add strInam, strTemp
					End If
				End If
				binItem.Close()
				p_start = p_end + intSeparator + 2
			Loop Until p_start + 3>lngRequestSize
			Me.FormItem = Split(strFormItem, strSplit)
			Me.FileItem = Split(strFileItem, strSplit)
		End If
	End Sub

	Private Function GetTimeStr()
		lngTime = lngTime + 1
		GetTimeStr = strDate&Right("00000000"&lngTime, 8)
	End Function

	Private Function GetFerr(lngFsiz, strFext)
		Dim intFerr
		intFerr = 0
		If lngFsiz>Me.MaxSize And Me.MaxSize>0 Then
			If Error = 0 Or Error = 2 Then Error = Error + 1
			intFerr = intFerr + 1
		End If
		If InStr(1, LCase("/"&Me.FileType&"/"), LCase("/"&strFext&"/")) = 0 And Me.FileType<>"" Then
			If Error<2 Then Error = Error + 2
			intFerr = intFerr + 2
		End If
		GetFerr = intFerr
	End Function

	Function Save(Item, strFnam)
		Save = False
		If objForm.Exists(Item&"_From") Then
			Dim intFerr, strFext
			strFext = objForm(Item&"_Ext")
			intFerr = GetFerr(objForm(Item&"_Size"), strFext)
			If objForm.Exists(Item&"_Err") Then
				If intFerr = 0 Then
					objForm(Item&"_Err") = 0
				End If
			Else
				objForm.Add Item&"_Err", intFerr
			End If
			If intFerr<>0 Then Exit Function
			If VarType(strFnam) = 2 Then
				Select Case strFnam
					Case 0
						strFnam = GetTimeStr()
						If strFext<>"" Then strFnam = strFnam&"."&strFext
					Case 1
						strFnam = objForm(Item&"_Name")
				End Select
			End If
			If Me.Mode = 1 Then
				objFormASP.Files(Item).SaveAS Server.MapPath(Me.SavePath&strFnam)
				objForm(Item&"_Path") = objFormASP.Files(Item).Path
			Else
				binForm.Position = objForm(Item&"_From")
				binItem.Type = 1
				binItem.Open
				binForm.CopyTo binItem, objForm(Item&"_Size")
				binItem.SaveToFile Server.MapPath(Me.SavePath&strFnam), 2
				binItem.Close()
			End If
			If objForm.Exists(Item) Then
				objForm(Item) = strFnam
			Else
				objForm.Add Item, strFnam
			End If
			Save = true
		End If
	End Function

	Function GetData(Item)
		GetData = ""
		If Me.Mode = 0 And objForm.Exists(Item&"_From") Then
			If GetFerr(objForm(Item&"_Size"), objForm(Item&"_Ext"))<>0 Then Exit Function
			binForm.Position = objForm(Item&"_From")
			GetData = binForm.Read(objForm(Item&"_Size"))
		End If
	End Function

	Function Form(Item)
		If objForm.Exists(Item) Then
			Form = objForm(Item)
		Else
			Form = ""
		End If
	End Function

	Private Function BinVal2(bin)
		Dim lngValue, i
		lngValue = 0
		For i = lenb(bin) To 1 step -1
			lngValue = lngValue * 256 + ascb(midb(bin, i, 1))
		Next
		BinVal2 = lngValue
	End Function

	Private Function Bin2Val(bin)
		Dim lngValue, i
		lngValue = 0
		For i = 1 To lenb(bin)
			lngValue = lngValue * 256 + ascb(midb(bin, i, 1))
		Next
		Bin2Val = lngValue
	End Function

End Class
%>