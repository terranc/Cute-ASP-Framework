<%
'**********
'	class		: TdbToExcel
'	File Name	: TdbToExcel.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-5-17
'**********


'**********
'	示例
'**********

'********** 

'**********
'	构建类
'**********

Class Class_Export
	Private s_table, s_fields, s_fieldnames,oConn
	Public FileName		'设置生成的文件名

	Public Property Let Conn(connObj)
		Set oConn = connObj
	End Property
	
	Private Sub Class_Initialize()
		FileName = "data"
	End Sub
	
	Public Default Property Get Contents(sTable,sFields,sWhere,sOrders,sAnother)
		Call Excel(sTable,sFields,sWhere,sOrders,sAnother)
	End Property
	
	Function Excel(sSql,sFields)
		Dim rs,aAnother,i
		Response.Buffer = True 
		Response.ContentType = "application/vnd.ms-excel" 
		Response.AddHeader "content-disposition", "inline; filename = "&Me.FileName&".xls"
		Response.Write "<table border=""1"">"
		Set rs = oConn.Execute(sSql)
		Response.Write("<tr>")
		If sFields = "" Then
			For i=0 To rs.fields.count-1
				Response.Write "<td>"&rs.Fields(i).Name&"</td>"
			Next
		Else
			aFields = Split(sFields,"|")
			For i=0 To uBound(aFields)
				Response.Write "<td>"&aFields(i)&"</td>"
			Next
		End If
		Response.Write("</tr>")
		Do While Not rs.eof
			Response.Write "<tr>"
			for i=0 to rs.fields.count-1  
				If Len(rs(i))>12 And IsNumeric(rs(i)) Then
					Response.Write "<td>@"&rs(i)&"</td>"
				Else
					Response.Write "<td>"&rs(i)&"</td>"
				End If
			Next
			Response.Write "</tr>"
			rs.MoveNext
		Loop
		rs.close : Set rs = Nothing
		Response.Write "</table>"
	End Function

	Function Txt(sSql,sFields)
		Dim rs,aAnother,i
		Response.Buffer = True 
		Response.ContentType = "application/octet-stream"
		Response.AddHeader "content-disposition", "attachment; filename = "&Me.FileName&".txt"
		Set rs = oConn.Execute(sSql)
		If sFields = "" Then
			For i=0 To rs.fields.count-1
				Response.Write rs.Fields(i).Name&Chr(9)
			Next
		Else
			aFields = Split(sFields,"|")
			For i=0 To uBound(aFields)
				Response.Write aFields(i)&Chr(9)
			Next
		End If
		Response.Write vbCrlf
		Do While Not rs.eof
			for i=0 to rs.fields.count-1  
				If Len(rs(i))>12 And IsNumeric(rs(i)) Then
					Response.Write "@"&Replace(rs(i),Chr(9)," ")&Chr(9)
				Else
					Response.Write Replace(rs(i),Chr(9)," ")&Chr(9)
				End If
			Next
			Response.Write vbCrlf
			rs.MoveNext
		Loop
		rs.close : Set rs = Nothing
	End Function
End Class
%> 
