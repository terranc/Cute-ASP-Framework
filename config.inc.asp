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
Casp.WebConfig("CodePage")		=	936				'设置站点编码
Casp.WebConfig("Charset")		=	"gb2312"		'设置站点字符集
Casp.WebConfig("FilterWord")	=	""				'设置过滤字符

Casp.db.ConnectionType = "ACCESS"
Casp.db.ServerIp = "localhost"
Casp.db.Database = "\aspframework\db.mdb"
Casp.db.UserName = "sa"
Casp.db.Password = ""

On Error Resume Next
Session.CodePage = Casp.WebConfig("CodePage")
Response.Charset = Casp.WebConfig("Charset")
Casp.Cookie.Mark = "cute_"		'设置Cookie名称前缀
Casp.Cache.Mark = "cute_"		'设置缓存名称前缀
Casp.Ubb.Mode = 0				'使用基本UBB
Casp.Date.TimeZone = 8			'设置所在时区
On Error Goto 0

Sub Finish()
    Dim RunTime : RunTime = round((Timer() - StartTime), 3)
    echo "<p style=""text-align:center;font-size:11px;"">Page rendered in " & RunTime & " seconds.</p>"
End Sub
%>
