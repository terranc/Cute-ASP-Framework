<%
'Option Explicit
Response.Buffer = True
Dim StartTime : StartTime = timer()
%>
<!--#include file="include/ext.asp"-->
<!--#include file="include/class.asp"-->
<!--#include file="include/library/db.asp"-->
<!--#include file="include/library/date.asp"-->
<!--#include file="include/library/array.asp"-->
<!--#include file="include/library/string.asp"-->
<!--#include file="include/library/params.asp"-->
<!--#include file="include/library/session.asp"-->
<%
Casp.WebConfig("CodePage")		=	936				'����վ�����
Casp.WebConfig("Charset")		=	"gb2312"		'����վ���ַ���
Casp.WebConfig("FilterWord")	=	""				'���ù����ַ�

Casp.db.ConnectionType = "ACCESS"
Casp.db.ServerIp = "localhost"
Casp.db.Database = "\aspframework\db.mdb"
Casp.db.UserName = "sa"
Casp.db.Password = ""

On Error Resume Next
Session.CodePage = Casp.WebConfig("CodePage")
Response.Charset = Casp.WebConfig("Charset")
Casp.Cookie.Mark = "cute_"		'����Cookie����ǰ׺
Casp.Cache.Mark = "cute_"		'���û�������ǰ׺
Casp.Ubb.Mode = 0				'ʹ�û���UBB
Casp.Date.TimeZone = 8			'��������ʱ��
On Error Goto 0

Sub Finish()
    Dim RunTime : RunTime = round((Timer() - StartTime), 3)
    echo "<p style=""text-align:center;font-size:11px;"">Page rendered in " & RunTime & " seconds.</p>"
End Sub
%>
