<%
'**********
'	class		: ubb code class
'	File Name	: ubb.asp
'	Version		: 0.1.0
'	Author		: TerranC
'	Date		: 2008-12-6
'**********

'**********
'	示例
'**********

'**********
'	构建类
'**********
Class Class_Ubb
	
	'**********
	'属性：0、1、2、3几个等级（1：支持文字样式；2：支持多媒体；3：支持动画脚本）
	'**********
	Public Mode

    Private Sub Class_Initialize()
        Mode = 0
    End Sub

	Public Function Encode(ByVal ReStr)
		If Len(ReStr)>0 Then
			ReStr = Replace(ReStr, " ", "&nbsp;")
		Else
			Exit Function
		End If
		Dim re, i
		Set re = New regexp
		re.IgnoreCase = true
		re.Global = True

		re.Pattern = "(js):"
		ReStr = re.Replace(ReStr, "j#115;:")
		re.Pattern = "(vbs):"
		ReStr = re.Replace(ReStr, "vb&#115;:")
		re.Pattern = "(script)"
		ReStr = re.Replace(ReStr, "&#115cript")
		re.Pattern = "(value)"
		ReStr = re.Replace(ReStr, "&#118alue")
		re.Pattern = "(document.cookie)"
		ReStr = re.Replace(ReStr, "documents&#46cookie")
		re.Pattern = "(on(mouse|exit|error|click|key))"
		ReStr = re.Replace(ReStr, "&#111n$2")

		re.Pattern = "(\[img\])(\S+?)(\[\/img\])"
		ReStr = re.Replace(ReStr, "<img src=""$2"" alt="""">")
		re.Pattern = "(\[URL\])(\S+?)(\[\/URL\])"
		ReStr = re.Replace(ReStr, "<a href=""$2"" target=""_blank"">$2</a>")
		re.Pattern = "(\[URL=(\S+?)\])(.+?)(\[\/URL\])"
		ReStr = re.Replace(ReStr, "<a href=""$2"" target=""_blank"">$3</a>")
		re.Pattern = "(\[email\])(\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+){0,100})(\[\/email\])"
		ReStr = re.Replace(ReStr, "<a href=""mailto:$2"">$2</a>")
		re.Pattern = "(\[email=(\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+){0,80})\])(.[^\[]*)(\[\/email\])"
		ReStr = re.Replace(ReStr, "<a href=""mailto:$2"" target=""_blank"">$6</a>")

		If Me.Mode >= 1 Then
			re.Pattern = "(\[b\])((.|\n){0,}?)(\[\/b\])"
			ReStr = re.Replace(ReStr, "<strong>$2</strong>")
			re.Pattern = "(\[i\])((.|\n){0,}?)(\[\/i\])"
			ReStr = re.Replace(ReStr, "<i>$2</i>")
			re.Pattern = "(\[u\])((.|\n){0,}?)(\[\/u\])"
			ReStr = re.Replace(ReStr, "<u>$2</u>")
			re.Pattern = "(\[code\])((.|\n){0,}?)(\[\/code\])"
			ReStr = re.Replace(ReStr, "<code class=""ubb""><p>$2</p></code>")
			re.Pattern = "(\[size=(\d+)\])((.|\n){0,}?)(\[\/size\])"
			ReStr = re.Replace(ReStr, "<font size=""$2"">$3</span>")
			re.Pattern = "\[h(\d+)\]((.|\n){0,}?)(\[\/h(\d+)\])"
			ReStr = re.Replace(ReStr, "<h$1>$2</h$1>")
			re.Pattern = "\[align=(center|left|right)\]((.|\n){0,}?)(\[\/align\])"
			ReStr = re.Replace(ReStr, "<div style=""display:block;text-align:$1;"">$2</div>")
			re.Pattern = "(\[quote\])"
			ReStr = re.Replace(ReStr, "<blockquote class=""ubb"">")
			re.Pattern = "(\[\/quote\])"
			ReStr = re.Replace(ReStr, "</blockquote>")
		End If

		If Me.Mode >= 2 Then
			re.Pattern = "\[(rm|mp|qt)=.+?\]"
			ReStr = re.Replace(ReStr, "[media]")
			re.Pattern = "\[\/(rm|mp|qt)=.+?\]"
			ReStr = re.Replace(ReStr, "[/media]")
			re.Pattern = "\[(dir|flash)=.+?\]"
			ReStr = re.Replace(ReStr, "[swf]")
			re.Pattern = "\[\/(dir|flash)=.+?\]"
			ReStr = re.Replace(ReStr, "[/swf]")
			re.Pattern = "(\[swf\])((http|https|ftp):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\+\-~`@\':!%#]|(&amp;)|&)+\.swf)(\[\/swf\])"
			ReStr = re.Replace(ReStr, "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" width=""352"" height=""288"" align=""middle""><param name=""movie"" value=""$2"" /><param name=""quality"" value=""high""><param name=""menu"" value=""false""><embed src=""$2"" quality=""high"" width=""352"" height=""288"" align=""middle"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" /></object>")

			re.Pattern = "(\[flash\])((http|https|ftp):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\+\-~`@\':!%#]|(&amp;)|&)+\.swf)(\[\/flash\])"
			ReStr = re.Replace(ReStr, "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" width=""352"" height=""288"" align=""middle""><param name=""movie"" value=""$2"" /><param name=""quality"" value=""high""><param name=""menu"" value=""false""><embed src=""$2"" quality=""high"" width=""352"" height=""288"" align=""middle"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" /></object>")
			re.Pattern = "\[media\]((http|ftp|mms|https):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\+\-~`@\':!%#]|(&amp;)|&)+\.(wmv|asf|wm|wma|wmv|wmx|wmd|avi|mpeg|mpg|mpa|mpe|dat|w1v|mp2|asx))\[\/media]"
			ReStr = re.Replace(ReStr, "<object><embed src=""$1"" autostart=""false"" playcount=""1""/></object>")

			re.Pattern = "\[media\]((http|ftp|mms|https):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\+\-~`@\':!%#]|(&amp;)|&)+\.(mp3|mid))\[\/media]"
			ReStr = re.Replace(ReStr, "<object classid=""CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" width=""352"" height=""45""><param name=""filename"" value=""$1""/><embed src=""$1"" playcount=""1""></embed></object>")

			re.Pattern = "\[media\]((http|ftp|rtsp|https):(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\+\-~`@\':!%#]|(&amp;)|&)+\.(ra|rm|rmj|rms|mnd|ram|rmm|r1m|rom|mns))\[\/media]"
			ReStr = re.Replace(ReStr, "<object classid=clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa width=""352"" height=""288""><param name=""src"" value=""$1""/><param name=""console"" value=""clip1""/><param name=""controls"" value=""imagewindow""/><param name=""autostart"" value=""false""/></object><object classid=""clsid:cfcdaa03-8be4-11cf-b84b-0020afbbccfa"" height=""32"" width=""352""><param name=""src"" value=""$1""/><param name=""controls"" value=""controlpanel""/><param name=""console"" value=""clip1""/></object>")
		End If

		If Me.Mode >= 3 Then
			re.Pattern = "(\[fly\])((.|\n){0,}?)(\[\/fly\])"
			ReStr = re.Replace(ReStr, "<marquee width=""100%"" behavior=""alternate"" scrollamount=""3"">$2</marquee>")
			re.Pattern = "(\[move\])((.|\n){0,}?)(\[\/move\])"
			ReStr = re.Replace(ReStr, "<marquee scrollamount=""3"">$2</marquee>")
			re.Pattern = "(\[color=(.{3,10})\])((.|\n){0,}?)(\[\/color\])"
			ReStr = re.Replace(ReStr, "<span style=""color:$2;"">$3</span>")
			re.Pattern = "(\[glow=.+?\])((.|\n)+?)(\[\/glow\])"
			ReStr = re.Replace(ReStr, "$2")
			re.Pattern = "(\[shadow=.+?\])((.|\n)+?)(\[\/shadow\])"
			ReStr = re.Replace(ReStr, "$2")
		End If

		If InStr(LCase(ReStr), "http://")>0 Then
			re.Pattern = "(^|[^<=""'])(http:(\/\/|\\\\)(([\w\/\\\+\-~`@:%])+\.)+([\w\/\\\.\=\?\+\-~`@\':!%#]|(&amp;)|&)+)"
			ReStr = re.Replace(ReStr, "$1<a target=""_blank"" href=""$2"">$2</a>")
		End If
		'识别www等开头的网址
		If InStr(LCase(ReStr), "www.")>0 Then
			re.Pattern = "(^|[^\/\\\w<=""])((www)\.(\w)+\.([\w\/\\\.\=\?\+\-~`@\'!%#]|(&amp;))+)"
			ReStr = re.Replace(ReStr, "$1<a target=""_blank"" href=""http://$2"">$2</a>")
		End If

		ReStr = Replace(ReStr, Chr(13)&Chr(10), "<br />")

		Set re = Nothing
		Encode = ReStr
	End Function
End Class
%>