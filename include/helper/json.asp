<%
'**********************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003 - This file is part of GAB_LIBRARY
'* For license refer to the license.txt
'***********************************************************************************************

'****************************************************************************************

'' @CLASSTITLE:        JSON
'' @CREATOR:        Michal Gabrukiewicz (gabru at grafix.at), Michael Rebec
'' @CONTRIBUTORS:    - Cliff Pruitt (opensource at crayoncowboy.com)
''                    - Sylvain Lafontaine
'' @CREATEDON:        2007-04-26 12:46
'' @CDESCRIPTION:    Comes up with functionality for JSON (http://json.org) to use within ASP.
''                     Correct escaping of characters, generating JSON Grammer out of ASP datatypes and structures
'' @REQUIRES:        -
'' @OPTIONEXPLICIT:    yes
'' @VERSION:        1.4

'****************************************************************************************

Class Class_JSON

    'private members
    Private output, innerCall

    'public members
    Public toResponse ''[bool] should generated results be directly written to the response? default = false

    '*********************************************************************************
    '* constructor
    '*********************************************************************************
    Public Sub class_initialize()
        newGeneration()
        toResponse = false
    End Sub

    '******************************************************************************************
    '' @SDESCRIPTION:    STATIC! takes a given string and makes it JSON valid
    '' @DESCRIPTION:    all characters which needs to be escaped are beeing replaced by their
    ''                    unicode representation according to the
    ''                    RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
    '' @PARAM:            val [string]: value which should be escaped
    '' @RETURN:            [string] JSON valid string
    '' asc 函数被替换成ascw函数以便支持中文
    '******************************************************************************************
    Public Function escape(Val)
        Dim cDoubleQuote, cRevSolidus, cSolidus
        cDoubleQuote = &h22
        cRevSolidus = &h5C
        cSolidus = &h2F

        Dim i, currentDigit
        For i = 1 To (Len(Val))
            currentDigit = Mid(Val, i, 1)
            If ascw(currentDigit) > &h00 And ascw(currentDigit) < &h1F Then
                currentDigit = escapequence(currentDigit)
            ElseIf ascw(currentDigit) >= &hC280 And ascw(currentDigit) <= &hC2BF Then
                currentDigit = "\u00" + Right(padLeft(Hex(Asc(currentDigit) - &hC200), 2, 0), 2)
            ElseIf ascw(currentDigit) >= &hC380 And ascw(currentDigit) <= &hC3BF Then
                currentDigit = "\u00" + Right(padLeft(Hex(ascw(currentDigit) - &hC2C0), 2, 0), 2)
            Else
                Select Case ascw(currentDigit)
                    Case cDoubleQuote
                         currentDigit = escapequence(currentDigit)
                    Case cRevSolidus
                         currentDigit = escapequence(currentDigit)
                    Case cSolidus
                         currentDigit = escapequence(currentDigit)
                End Select
            End If
            escape = escape & currentDigit
        Next
    End Function

    '******************************************************************************
    '' @SDESCRIPTION:    generates a representation of a name value pair in JSON grammer
    '' @DESCRIPTION:    the generation is done fully recursive so the value can be a complex datatype as well. e.g.
    ''                    toJSON("n", array(array(), 2, true), false) or toJSON("n", array(RS, dict, false), false), etc.
    '' @PARAM:            name [string]: name of the value (accessible with javascript afterwards). leave empty to get just the value
    '' @PARAM:            val [variant], [int], [float], [array], [object], [dictionary], [recordset]: value which needs
    ''                    to be generated. Conversation of the data types (ASP datatype -> Javascript datatype):
    ''                    NOTHING, NULL -> null, ARRAY -> array, BOOL -> bool, OBJECT -> name of the type,
    ''                    MULTIDIMENSIONAL ARRAY -> generates a 1 dimensional array (flat) with all values of the multidim array
    ''                    DICTIONARY -> valuepairs. each key is accessible as property afterwards
    ''                    RECORDSET -> array where each row of the recordset represents a field in the array.
    ''                    fields have properties named after the column names of the recordset (LOWERCASED!)
    ''                    e.g. generate(RS) can be used afterwards r[0].ID
    ''                    INT, FLOAT -> number
    ''                    OBJECT with reflect() method -> returned as object which can be used within JavaScript
    '' @PARAM:            nested [bool]: is the value pair already nested within another? if yes then the {} are left out.
    '' @RETURN:            [string] returns a JSON representation of the given name value pair
    ''                    (if toResponse is on then the return is written directly to the response and nothing is returned)
    '*******************************************************************************************
    Public Function toJSON(Name, Val, nested)
        If Not nested And Not IsNull(Name) Then Write("{")
        If Not IsNull(Name) Then Write("""" & escape(Name) & """: ")
        generateValue(Val)
        If Not nested And Not IsNull(Name) Then Write("}")
        toJSON = output

        If innerCall = 0 Then newGeneration()
    End Function

    '*********************************************************************************
    '* generate
    '******************************************************************************
    Private Function generateValue(Val)
        If IsNull(Val) Then
            Write("null")
        ElseIf IsArray(Val) Then
            generateArray(Val)
        ElseIf IsObject(Val) Then
            If Val Is Nothing Then
                Write("null")
            ElseIf TypeName(Val) = "Dictionary" Then
                generateDictionary(Val)
            ElseIf TypeName(Val) = "Recordset" Then
                generateRecordset(Val)
            Else
                generateObject(Val)
            End If
        Else
            'bool
            varTyp = VarType(Val)
            If varTyp = 11 Then
                If Val Then Write("true") Else Write("false")
                'int, long, byte
            ElseIf varTyp = 2 Or varTyp = 3 Or varTyp = 17 Or varTyp = 19 Then
                Write(CLng(Val))
                'single, double, currency
            ElseIf varTyp = 4 Or varTyp = 5 Or varTyp = 6 Or varTyp = 14 Then
                Write(Replace(CDbl(Val), ",", "."))
            Else
                Write("""" & escape(Val & "") & """")
            End If
        End If
        generateValue = output
    End Function

    '*****************************************************************************
    '* generateArray
    '*****************************************************************************
    Private Sub generateArray(Val)
        Dim Item, i
        Write("[")
        i = 0
        'the for each allows us to support also multi dimensional arrays
        For Each Item in Val
            If i > 0 Then Write(",")
            generateValue(Item)
            i = i + 1
        Next
        Write("]")
    End Sub

    '*********************************************************************************
    '* generateDictionary
    '**************************************************************************
    Private Sub generateDictionary(Val)
        Dim Keys, i
        innerCall = innerCall + 1
        Write("{")
        Keys = Val.Keys
        For i = 0 To UBound(Keys)
            If i > 0 Then Write(",")
            toJSON Keys(i), Val(Keys(i)), true
        Next
        Write("}")
        innerCall = innerCall - 1
    End Sub

    '*******************************************************************
    '* generateRecordset
    '*******************************************************************
    Private Sub generateRecordset(Val)
        Dim i
        Write("[")
        While Not Val.EOF
            innerCall = innerCall + 1
            Write("{")
            For i = 0 To Val.fields.Count - 1
                If i > 0 Then Write(",")
                toJSON LCase(Val.fields(i).Name), Val.fields(i).Value, true
            Next
            Write("}")
            Val.movenext()
            If Not Val.EOF Then Write(",")
            innerCall = innerCall - 1
        Wend
        Write("]")
    End Sub

    '*******************************************************************************
    '* generateObject
    '*******************************************************************************
    Private Sub generateObject(Val)
        Dim props
        On Error Resume Next
        Set props = Val.reflect()
        If Err = 0 Then
            On Error GoTo 0
            innerCall = innerCall + 1
            toJSON Empty, props, true
            innerCall = innerCall - 1
        Else
            On Error GoTo 0
            Write("""" & escape(TypeName(Val)) & """")
        End If
    End Sub

    '*******************************************************************************
    '* newGeneration
    '*******************************************************************************
    Private Sub newGeneration()
        output = Empty
        innerCall = 0
    End Sub

    '*******************************************************************************
    '* JsonEscapeSquence
    '*******************************************************************************
    Private Function escapequence(digit)
        escapequence = "\u00" + Right(padLeft(Hex(Asc(digit)), 2, 0), 2)
    End Function

    '*****************************************************************************
    '* padLeft
    '*****************************************************************************
    Private Function padLeft(Value, totalLength, paddingChar)
        padLeft = Right(clone(paddingChar, totalLength) & Value, totalLength)
    End Function

    '*****************************************************************************
    '* clone
    '******************************************************************************************
    Public Function clone(byVal Str, n)
        Dim i
        For i = 1 To n
             clone = clone & Str
        Next
    End Function

	'******************************************************************************************
	'* write
	'******************************************************************************************
	Private Sub Write(Val)
		If toResponse Then
			response.Write(Val)
		Else
			output = output & Val
		End If
	End Sub

End Class
%>
