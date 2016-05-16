<!--#include file="config.inc.asp"-->
<!--#include file="interface/demo.asp"-->
<%
Set oParams = Server.CreateObject("Scripting.Dictionary")
oParams.CompareMode = 1
For Each Item In Request.QueryString()
	If LCase(Item) <> "module" And LCase(Item) <> "action" Then
		oParams.Add Item,Tpub.rq(2,Item,1,"")
	End If
Next
For Each Item In Request.Form()
	If LCase(Item) <> "module" And LCase(Item) <> "action" Then
		oParams.Add Item,Tpub.rq(3,Item,1,"")
	End If
Next
On Error Resume Next
ExecuteGlobal("call "&Request("action")&"(oParams)")
Set oParams = Nothing
If Err Then
	'die Err.Description
	die "-999"
End If
On Error Goto 0
%>
