<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI回傳檔下載</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.btn3{
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">DCI回傳檔下載列表<img src="space.gif" width="15" height="8"></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">時間</th>
					<th height="34">檔案</th>
					<th height="34">下載</th>
				</tr><%
						nowTime=request("nowTime")
						FileType=".txt"
						reUpFileName=""
						Set Fso=CreateObject("Scripting.FileSystemObject")
						
						For i = 3 to 0 step -1
							
							fileyear=year(dateadd("m",-(i),now))-1911
							fileMonth=right("00"&month(dateadd("m",-(i),now)),2)

							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">" 
							Response.Write "<td>"
							Response.Write fileyear&"年"&fileMonth&"月"
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write fileyear&fileMonth&".zip"
							Response.Write "</td>"

							Response.Write "<td>"
							If Fso.FileExists(server.mappath("/traffic/dcifile/"&fileyear&fileMonth&".zip")) then
								
								response.write "<input type=""button"" name=""ReSend"" value=""下載"" class=""btn3"" style=""width:40px;height:25px;"" onclick=""location.href='/traffic/dcifile/"&fileyear&fileMonth&".zip';"">"
							else
								Response.Write "暫無檔案"
							end if
							Response.Write "</td>"

							Response.Write "</tr>"
						Next
					%>
				<tr bgcolor="#ffffff" align="center">
					<td height="35" bgcolor="#1BF5FF" colspan="3">
						<input type="button" name="exit" value=" 離 開 " onclick="funExt();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<script language="javascript">
	window.resizeTo(800,400)
	function funExt() {
		self.close();
	}
</script>
