<%
'**********
'	class		: file operate class
'	File Name	: file.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-5-17
'**********

'**********
'	ʾ��
'**********

'**********
'	������
'**********

Class Class_File
    Public FSO '����FSO�������
    Public Stream '����Stream�������
    Public Charset '�����ַ���

    '**********
    ' ������: class_Initialize
    ' ��  ��: Constructor
    '**********

    Private Sub Class_Initialize()
        Charset = "gb2312"
        Stream = "ADODB.Stream"
        FSO = "Scripting.FileSystemObject"
    End Sub

    '**********
    '��������GetFileTypeName
    '��  �ã���ȡ�ļ���չ��
    '**********
    Function GetFileTypeName(fName)
        Dim fileName
        fileName = Split(fName, ".")
        GetFileTypeName = fileName(UBound(fileName))
    End Function

    '**********
    '��������CheckFileExt
    '��  �ã������Ƿ�Ϸ��ļ�
    '**********
    Function CheckFileExt(sFile, ext)
        If Trim(ext) = "" And Trim(ext) <> "*" Then
            CheckFileExt = "jpg|gif|png"
        End If
        Dim arr_Ext
        arr_Ext = Split(ext, "|")
        Dim i, fileExt
        fileExt = Me.GetFileTypeName(sFile)
        CheckFileExt = false
        For Each i In arr_Ext
            If fileExt = i Then CheckFileExt = true
        Next
    End Function

    '**********
    'FormatFileSize
    '��  �ã���ʽ���ļ��Ĵ�С
    '**********
    Function FormatFileSize(fs,float)
        Dim bUnit,kUnit, mUnit, gUnit
        bUnit = "B"
        kUnit = "KB"
        mUnit = "MB"
        gUnit = "GB"
		If Not IsNumeric(float) Then float = 0
        If fs>1073741824 Then
            fs = FormatNumber(fs / 1073741824, float, true)&gUnit
        ElseIf fs>1048576 Then
            fs = FormatNumber(fs / 1048576, float, true)&mUnit
        ElseIf fs>1024 Then
            fs = FormatNumber(fs / 1024, float, true)&kUnit
        Else
            fs = FormatNumber(fs, float, true)&bUnit
        End If
        FormatFileSize = fs
    End Function

    '**********
    'GetFileSize
    '��  �ã���ȡ�ļ��Ĵ�С
    '**********
    Function GetFileSize(fls)
        Dim fso, fdr, arr_fls, fsize, i, fl
        arr_fls = Split(fls, "||")
        fsize = 0
        For Each i In arr_fls
            fl = Server.MapPath(i)
            Set fso = Server.CreateObject(Me.FSO)
            If fso.FileExists(fl) Then
                Set fdr = fso.GetFile(fl)
                fsize = fdr.Size + fsize
                Set fdr = Nothing
            End If
            Set fso = Nothing
        Next
        GetFileSize = fsize
    End Function

    '**********
    'GetFolderSize
    '��  �ã���ȡĿ¼�Ĵ�С
    '**********
    Function GetFolderSize(fls)
        Dim fso, fdr, arr_fls, fsize, i, fl
        arr_fls = Split(fls, "||")
        fsize = 0
        For Each i In arr_fls
            fl = Server.MapPath(i)
            Set fso = Server.CreateObject(Me.FSO)
            If fso.FolderExists(fl) Then
                Set fdr = fso.GetFolder(fl)
                fsize = fdr.Size + fsize
                Set fdr = Nothing
            End If
            Set fso = Nothing
        Next
        GetFolderSize = fsize
    End Function

    '**********
    '��������IsFolderExists
    '��  �ã����ĳһĿ¼�Ƿ����
    '��  ����FolderPath	----Ŀ¼
    '**********
    Function IsFolderExists(FolderPath)
        Dim fso
        FolderPath = Server.MapPath(".")&"\"&FolderPath
        Set fso = Server.CreateObject(Me.FSO)
        If fso.FolderExists(FolderPath) Then
            IsFolderExists = true '����
        Else
            IsFolderExists = false '������
        End If
        Set fso = Nothing
    End Function

    '**********
    '��������IsFileExists
    '��  �ã����ĳһ�ļ��Ƿ����
    '��  ����FilePath	----Ŀ¼
    '**********
    Function IsFileExists(FilePath)
        Dim fso
        FilePath = Server.MapPath(".")&"\"&FilePath
        Set fso = Server.CreateObject(Me.FSO)
        If fso.FileExists(FilePath) Then
            IsFileExists = true '����
        Else
            IsFileExists = false '������
        End If
        Set fso = Nothing
    End Function

    '**********
    '��������CreatePath
    '��  �ã������༶Ŀ¼�����Դ��������ڵĸ�Ŀ¼
    '��  ����Ҫ������Ŀ¼���ƣ������Ƕ༶
    '�����߼�ֵ��True�ɹ���Falseʧ��
    '����Ŀ¼�ĸ�Ŀ¼�ӵ�ǰĿ¼��ʼ
    '**********
    Function CreatePath(CFolder)
        On Error Resume Next
        Dim objFSO, PhCreateFolder, CreateFolderArray, CreateFolder
        Dim i, ii, CreateFolderSub, PhCreateFolderSub, BlInfo
        BlInfo = false
        CreateFolder = CFolder
        Set objFSO = Server.CreateObject(Me.FSO)
        If Err Then
            Err.Clear
            Exit Function
        End If
        CreateFolder = Replace(CreateFolder, "\", "/")
        If Right(CreateFolder, 1) = "/" Then
            CreateFolder = Left(CreateFolder, Len(CreateFolder) -1)
        End If
        CreateFolderArray = Split(CreateFolder, "/")
        For i = 0 To UBound(CreateFolderArray)
            CreateFolderSub = ""
            For ii = 0 To i
                CreateFolderSub = CreateFolderSub & CreateFolderArray(ii) & "/"
            Next
            PhCreateFolderSub = Server.MapPath(CreateFolderSub)
            If Not objFSO.FolderExists(PhCreateFolderSub) Then
                objFSO.CreateFolder(PhCreateFolderSub)
            End If
        Next
        If Err Then
            Err.Clear
        Else
            BlInfo = true
        End If
        Set objFSO = Nothing
        CreatePath = BlInfo
    End Function

    '**********
    '��������DelFolder
    '��  �ã�ɾ��Ŀ¼
    '��  ����Ҫɾ����Ŀ¼����
    '**********
    Function DelFolder(sPath)
        On Error Resume Next
        DelFolder = false
        Dim fso, tmpfolder, tmpsubfolder, tmpfile, tmpfiles
        Set fso = Server.CreateObject(Me.FSO)
        If (fso.FolderExists(Server.MapPath (sPath))) Then
            Set tmpfolder = fso.GetFolder(Server.MapPath (sPath))
            Set tmpfiles = tmpfolder.Files
            For Each tmpfile in tmpfiles
                fso.DeleteFile (tmpfile)
            Next
            Set tmpsubfolder = tmpfolder.SubFolders
            For Each tmpfolder in tmpsubfolder
                DelFolder(spath&"/"&tmpfolder.Name )
            Next
            fso.DeleteFolder (Server.MapPath (sPath))
        End If
        If Err Then
            Err.Clear
        Else
            DelFolder = true
        End If
    End Function

    '**********
    '��������DelFile
    '��  �ã�ɾ���ļ�
    '��  ����Ҫɾ�����ļ�����(֧����|�ָ����б�)
    '**********
    Function DelFile(sFiles)
        DelFile = true
        Dim fso, sFile, i
        sFile = Split(sFiles, "|")
        Set fso = Server.CreateObject(Me.FSO)
        For i = 0 To UBound(sFile)
            If fso.FileExists(Server.MapPath(sFile(i))) Then
                fso.DeleteFile(Server.MapPath(sFile(i)))
            End If
        Next
        Set fso = Nothing
        If Err Then
            Err.Clear
            DelFile = false
        End If
    End Function

    '**********
    '��������LoadFile
    '��  �ã���ȡ�ļ�
    '��  ����File	----�ļ�·��
    '**********
    Function LoadFile(sFile)
        On Error Resume Next
        Dim objStream
        Dim RText
        Set objStream = Server.CreateObject(Me.Stream)
        If Err Then
            RText = Array(Err.Number, Err.Description)
            LoadFile = "False"
            Err.Clear
            Exit Function
        End If
        With objStream
            .Type = 2
            .Mode = 3
            .Open
            .Charset = Charset
            .Position = objStream.Size
            .LoadFromFile Server.MapPath(sFile)
            If Err.Number<>0 Then
                RText = Err.Description
                LoadFile = RText
                Err.Clear
                Exit Function
            End If
            RText = .ReadText
            .Close
        End With
        LoadFile = RText
        Set objStream = Nothing
    End Function

    '**********
    '��������SaveFile
    '��  �ã������ļ�
    '��  ����sFilePath	----�ļ�·��
    '		 sPageContent --�ļ�����
    '**********
    Function SaveFile(sFilePath, sPageContent)
        SaveFile = true
        Dim FileName
        Dim S
        Set S = Server.CreateObject(Me.Stream)
        FileName = Server.MapPath(sFilePath)
        With S
			If VarType(sPageContent) = 8209 Then
				.Type = 1
			End If
            .Open
			If VarType(sPageContent) = 8209 Then
				.Write sPageContent
			Else
			    .Charset = Charset
				.WriteText sPageContent
			End If
            .SaveToFile FileName, 2
            .Close
        End With
        Set S = Nothing
        If Err Then
            SaveFile = false
        End If
    End Function

    '**********
    '����Ŀ¼�������ļ�
    'sFolderPath:ԴĿ¼
    'dFolderPath:Ŀ��Ŀ¼
    '**********
    Function CopyFolder(sFolderPath, dFolderPath)
        On Error Resume Next
        CopyFolder = true
        Dim fs
        Set fs = Server.CreateObject(Me.FSO)
        fs.CopyFolder Server.Mappath(sFolderPath), Server.Mappath(dFolderPath)
        Set fs = Nothing
        If Err Then
            CopyFolder = false
        End If
    End Function

    '**********
    '�����ļ�
    'sFilePath:Դ�ļ�
    'dFilePath:Ŀ�ļ�
    '**********
    Function CopyFile(sFilePath, dFilePath)
        On Error Resume Next
        CopyFile = true
        Dim fs
        Set fs = Server.CreateObject(Me.FSO)
        fs.CopyFile Server.Mappath(sFilePath), Server.Mappath(dFilePath)
        Set fs = Nothing
        If Err Then
            Err.Clear
            CopyFile = false
        End If
    End Function

    '**********
    '����ָ��Ŀ¼�ļ��б�
    'strDir:Ŀ¼
    'strFileExt:�ļ�����(��|�ָ�)
    '**********
    Function LoadIncludeFiles(strDir, strFileExt)
        Dim aryFileList()
        ReDim aryFileList(0)
        Dim fso, f, f1, fc, s, i
        Set fso = Server.CreateObject(Me.FSO)
        Set f = fso.GetFolder(Server.Mappath(strDir))
        Set fc = f.Files
        i = 0
        For Each f1 in fc
            If Me.CheckFileExt(f1.Name, strFileExt) Then
                ReDim Preserve aryFileList(i)
                aryFileList(i) = f1.Name
                i = i + 1
            End If
        Next
        LoadIncludeFiles = aryFileList
    End Function

    '**********
    '����ָ��Ŀ¼��Ŀ¼�б�
    'strDir:Ŀ¼
    '**********
    Function LoadIncludeFolder(strDir)
        Dim aryFileList()
        ReDim aryFileList(0)
        Dim fso, f, f1, fc, s, i
        Set fso = Server.CreateObject(Me.FSO)
        Set f = fso.GetFolder(Server.Mappath(strDir))
        Set fc = f.SubFolders
        i = 0
        For Each f1 in fc
            ReDim Preserve aryFileList(i)
            aryFileList(i) = f1.Name
            i = i + 1
        Next
        LoadIncludeFolder = aryFileList
    End Function

    '**********
    'ȥ��utf-8ǩ��
    'sFilepath:�ļ�·��
    '**********
    Sub RemoveUtf8bom(sFilepath)
        On Error Resume Next
        Dim oFile, oStream, oXml, oStream2, oElement
        oFile = Server.MapPath(sFilepath)
        Set oStream = server.CreateObject(Me.Stream)
        With oStream
            .Type = 1
            .Open()
            .loadfromfile oFile
        End With
        Set oXml = server.CreateObject("Msxml2.DOMDocument.3.0")
        Set oElement = oXml.CreateElement("file")
        With oElement
            .DataType = "bin.base64"
            .NodeTypedValue = oStream.Read(3)
        End With
        If oElement.text = "77u/" Then
            oStream.Position = 3
            Set oStream2 = Server.CreateObject(Me.Stream)
            With oStream2
                .mode = 3
                .Type = 1
                .Open()
            End With
            oStream.CopyTo(oStream2)
            oStream2.SaveToFile oFile, 2
        End If
        Set oStream = Nothing
        Set oStream2 = Nothing
        Set oElement = Nothing
        Set oXml = Nothing
    End Sub
End Class
%>