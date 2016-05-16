<%
'**********
'	class		: Interface
'	File Name	: Class_Interface.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-6-27
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********
Class Class_XML
    Private fNode, fANode
    Private fErrInfo, fFileName, fOpen
    Public XmlDom

    '开打一个已经存在的XML文件,返回打开状态
    Public Function Open(byVal XmlSourceFile)
        Open = false
        fErrInfo = ""
        fFileName = ""
        fopen = false
        Set fNode = Nothing
        Set fANode = Nothing
        XmlSourceFile = Trim(XmlSourceFile)
        If XmlSourceFile = "" Then Exit Function
		On Error Resume Next
	    Set XmlDom = CreateObject("Msxml2.DOMDocument.3.0")
		If Err Then
			Err.Clear
			Set XmlDom = CreateObject("Microsoft.XMLDOM")
		End If
		On Error Goto 0
        XmlDom.preserveWhiteSpace = true
        XmlDom.async = False
        If Left(XmlSourceFile,5) = "<?xml" Then
	        XmlDom.loadXML XmlSourceFile
		Else
			XmlDom.load Server.MapPath(XmlSourceFile)
		End If

        fFileName = XmlSourceFile
        If Not IsError Then
            Open = true
            fopen = true
        End If
    End Function

    '关闭
    Public Sub Close()
		Set XmlDom = Nothing
        Set fNode = Nothing
        Set fANode = Nothing

        fErrInfo = ""
        fFileName = ""
        fopen = false
    End Sub

    '给xml内容
    Public Property Get XmlSource(byVal ElementOBJ)
        If fopen = false Then Exit Property

        Set ElementOBJ = ChildNode(XmlDom, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Set ElementOBJ = Nothing : Exit Property

        XmlSource = ElementOBJ.xml
    End Property

    '返回节点的缩进字串
    Private Property Get TabStr(byVal Node)
        TabStr = ""
        If Node Is Nothing Then Exit Property
        If Not Node.parentNode Is Nothing Then TabStr = "  "&TabStr(Node.parentNode)
    End Property

    '返回一个子节点对象,ElementOBJ为父节点,ChildNodeObj要查找的节点,IsAttributeNode指出是否为属性对象
    Public Property Get ChildNode(byVal ElementOBJ, byVal ChildNodeObj, byVal IsAttributeNode)
        Dim Element
        Set ChildNode = Nothing

        If IsNull(ChildNodeObj) Then
            If IsAttributeNode = false Then
                Set ChildNode = fNode
            Else
                Set ChildNode = fANode
            End If
            Exit Property
        ElseIf IsObject(ChildNodeObj) Then
            Set ChildNode = ChildNodeObj
            Exit Property
        End If

        Set Element = Nothing
        If LCase(TypeName(ChildNodeObj)) = "string" And Trim(ChildNodeObj)<>"" Then
            If IsNull(ElementOBJ) Then
                Set Element = fNode
            ElseIf LCase(TypeName(ElementOBJ)) = "string" Then
                If Trim(ElementOBJ)<>"" Then
                    Set Element = XmlDom.selectSingleNode("//"&Trim(ElementOBJ))
                    If LCase(Element.nodeTypeString) = "attribute" Then Set Element = Element.selectSingleNode("..")
                End If
            ElseIf IsObject(ElementOBJ) Then
                Set Element = ElementOBJ.selectSingleNode("..")
            End If

            If Element Is Nothing Then
                Set ChildNode = XmlDom.selectSingleNode("//"&Trim(ChildNodeObj))
            ElseIf IsAttributeNode = true Then
                Set ChildNode = Element.selectSingleNode("./@"&Trim(ChildNodeObj))
            Else
                Set ChildNode = Element.selectSingleNode("./"&Trim(ChildNodeObj))
            End If
        End If
    End Property

    '建立一个XML文件，RootElementName：根结点名。XSLURL：使用XSL样式地址
    '返回根结点
    Public Function Create(byVal RootElementName, byVal XslUrl)
        Dim PINode, RootElement

        Set Create = Nothing

        If (XmlDom Is Nothing) Or (fopen = true) Then Exit Function

        If Trim(RootElementName) = "" Then RootElementName = "Root"

        Set PINode = XmlDom.CreateProcessingInstruction("xml", "version=""1.0""  encoding=""utf-8""")
        XmlDom.appendChild PINode

        Set PINode = XMLDOM.CreateProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href="""&XslUrl&"""")
        XmlDom.appendChild PINode

        Set RootElement = XmlDom.createElement(Trim(RootElementName))
        XmlDom.appendChild RootElement

        Set Create = RootElement

        fopen = true
        Set fNode = RootElement
    End Function

    '插入在BefelementOBJ下面一个名为ElementName，Value为ElementText的子节点。
    'IsFirst：是否插在第一个位置；IsCDATA：说明节点的值是否属于CDATA类型
    '插入成功就返回新插入这个节点
    'BefelementOBJ可以是对象也可以是节点名，为空就取当前默认对象
    Public Function InsertElement(byVal BefelementOBJ, byVal ElementName, byVal ElementText, byVal IsFirst, byVal IsCDATA)
        Dim Element, TextSection, SpaceStr
        Set InsertElement = Nothing

        If Not fopen Then Exit Function

        Set BefelementOBJ = ChildNode(XmlDom, BefelementOBJ, false)
        If BefelementOBJ Is Nothing Then Exit Function

        Set Element = XmlDom.CreateElement(Trim(ElementName))
        If IsFirst = true Then
            BefelementOBJ.InsertBefore Element, BefelementOBJ.firstchild
        Else
            BefelementOBJ.appendChild Element
        End If

        If IsCDATA = true Then
            Set TextSection = XmlDom.createCDATASection(ElementText)
            Element.appendChild TextSection
        ElseIf ElementText<>"" Then
            Element.Text = ElementText
        End If

        Set InsertElement = Element
        Set fNode = Element
    End Function

    '在ElementOBJ节点上插入或修改名为AttributeName，值为：AttributeText的属性
    '如果已经存在名为AttributeName的属性对象，就进行修改。
    '返回插入或修改属性的Node
    'ElementOBJ可以是Element对象或名，为空就取当前默认对象
    Public Function SetAttributeNode(byVal ElementOBJ, byVal AttributeName, byVal AttributeText)
        Dim AttributeNode
        Set SetAttributeNode = Nothing

        If Not fopen Then Exit Function

        Set ElementOBJ = ChildNode(XmlDom, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Exit Function

        Set AttributeNode = ElementOBJ.Attributes.getNamedItem(AttributeName)
        If AttributeNode Is Nothing Then
            Set AttributeNode = XmlDom.CreateAttribute(AttributeName)
            ElementOBJ.SetAttributeNode AttributeNode
        End If
        AttributeNode.text = AttributeText

        Set fNode = ElementOBJ
        Set fANode = AttributeNode
        Set SetAttributeNode = AttributeNode
    End Function

    '修改ElementOBJ节点的Text值，并返回这个节点
    'ElementOBJ可以对象或对象名，为空就取当前默认对象
    Public Function UpdateNodeText(byVal ElementOBJ, byVal NewElementText, byVal IsCDATA)
        Dim TextSection

        Set UpdateNodeText = Nothing
        If Not fopen Then Exit Function

        Set ElementOBJ = ChildNode(XmlDom, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Exit Function

        If IsCDATA = true Then
            Set TextSection = XmlDom.createCDATASection(NewElementText)
            If ElementOBJ.firstchild Is Nothing Then
                ElementOBJ.appendChild TextSection
            ElseIf LCase(ElementOBJ.firstchild.nodeTypeString) = "cdatasection" Then
                ElementOBJ.replaceChild TextSection, ElementOBJ.firstchild
            End If
        Else
            ElementOBJ.Text = NewElementText
        End If

        Set fNode = ElementOBJ
        Set UpdateNodeText = ElementOBJ
    End Function

    '读取一个NodeOBJ的节点Text的值
    'NodeOBJ可以是节点对象或节点名，为空就取当前默认fNode
    Public Function GetNodeText(byVal NodeOBJ)
        GetNodeText = ""
        If fopen = false Then Exit Function

        Set NodeOBJ = ChildNode(null, NodeOBJ, false)
        If NodeOBJ Is Nothing Then Exit Function

        If LCase(NodeOBJ.nodeTypeString) = "element" Then
            Set fNode = NodeOBJ
        Else
            Set fANode = NodeOBJ
        End If
        GetNodeText = NodeOBJ.text
    End Function

    '返回符合testValue条件的第一个ElementNode，为空就取当前默认对象
    Public Function GetElementNode(byVal ElementName, byVal testValue)
        Dim Element, regEx, baseName

        Set GetElementNode = Nothing
        If Not fopen Then Exit Function

        testValue = Trim(testValue)
        Set regEx = New RegExp
        regEx.Pattern = "^[A-Za-z]+"
        regEx.IgnoreCase = true
        If regEx.Test(testValue) Then testValue = "/"&testValue
        Set regEx = Nothing

        baseName = LCase(Right(ElementName, Len(ElementName) - InStrRev(ElementName, "/", -1)))

        Set Element = XmlDom.SelectSingleNode("//"&ElementName&testValue)

        If Element Is Nothing Then
            Set GetElementNode = Nothing
            Exit Function
        End If

        Do While LCase(Element.baseName)<>baseName
            Set Element = Element.selectSingleNode("..")
            If Element Is Nothing Then Exit Do
        Loop

        If LCase(Element.baseName)<>baseName Then
            Set GetElementNode = Nothing
        Else
            Set GetElementNode = Element
            If LCase(Element.nodeTypeString) = "element" Then
                Set fNode = Element
            Else
                Set fANode = Element
            End If
        End If
    End Function

    '删除一个子节点
    Public Function RemoveChild(byVal ElementOBJ)
        RemoveChild = false
        If Not fopen Then Exit Function

        Set ElementOBJ = ChildNode(null, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Exit Function

        If LCase(ElementOBJ.nodeTypeString) = "element" Then
            If ElementOBJ Is fNode Then Set fNode = Nothing
            If ElementOBJ.parentNode Is Nothing Then
                XmlDom.RemoveChild(ElementOBJ)
            Else
                ElementOBJ.parentNode.RemoveChild(ElementOBJ)
            End If
            RemoveChild = true
        End If
    End Function

    '清空一个节点所有子节点
    Public Function ClearNode(byVal ElementOBJ)
        Set ClearNode = Nothing
        If Not fopen Then Exit Function

        Set ElementOBJ = ChildNode(null, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Exit Function

        ElementOBJ.text = ""
        ElementOBJ.RemoveChild(ElementOBJ.firstchild)

        Set ClearNode = ElementOBJ
        Set fNode = ElementOBJ
    End Function

    '删除子节点的一个属性
    Public Function RemoveAttributeNode(byVal ElementOBJ, byVal AttributeOBJ)
        RemoveAttributeNode = false
        If Not fopen Then Exit Function

        Set ElementOBJ = ChildNode(XmlDom, ElementOBJ, false)
        If ElementOBJ Is Nothing Then Exit Function

        Set AttributeOBJ = ChildNode(ElementOBJ, AttributeOBJ, true)
        If Not AttributeOBJ Is Nothing Then
            ElementOBJ.RemoveAttributeNode(AttributeOBJ)
            RemoveAttributeNode = true
        End If
    End Function

    '保存打开过的文件，只要保证FileName不为空就可以实现保存
    Public Function Save()
        On Error Resume Next
        Save = false
        If (Not fopen) Or (fFileName = "") Then Exit Function

        XmlDom.Save fFileName
        Save = (Not IsError)
        If Err.Number<>0 Then
            Err.Clear
            Save = false
        End If
    End Function

    '另存为XML文件，只要保证FileName不为空就可以实现保存
    Public Function SaveAs(SaveFileName)
        On Error Resume Next
        SaveAs = false
        If (Not fopen) Or SaveFileName = "" Then Exit Function
        XmlDom.Save SaveFileName
        SaveAs = (Not IsError)
        If Err.Number<>0 Then
            Err.Clear
            SaveAs = false
        End If
    End Function

    '检查并打印错误信息
    Private Function IsError()
        If XmlDom.ParseError.errorcode<>0 Then
            fErrInfo = "<h1>Error"&XmlDom.ParseError.errorcode&"</h1>"
            fErrInfo = fErrInfo&"<B>Reason :</B>"&XmlDom.ParseError.reason&"<br>"
            fErrInfo = fErrInfo&"<B>URL &nbsp; &nbsp;:</B>"&XmlDom.ParseError.url&"<br>"
            fErrInfo = fErrInfo&"<B>Line &nbsp; :</B>"&XmlDom.ParseError.Line&"<br>"
            fErrInfo = fErrInfo&"<B>FilePos:</B>"&XmlDom.ParseError.filepos&"<br>"
            fErrInfo = fErrInfo&"<B>srcText:</B>"&XmlDom.ParseError.srcText&"<br>"
            IsError = true
        Else
            IsError = false
        End If
    End Function

    '读取最后的错误信息
    Public Property Get ErrInfo
        ErrInfo = fErrInfo
    End Property
End Class
%>
