<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'fMnoth=month(now)
'if fMnoth<10 then
'fMnoth="0"&fMnoth
'end if
'fDay=day(now)
'if fDay<10 then
'fDay="0"&fDay
'end if
'fname=year(now)&fMnoth&fDay&"_批次文件.doc"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>批次輸出系統</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	line-height:2;
}
.style2 {font-size: 18px; font-family: "標楷體"; line-height:2;}
.style3 {font-size: 18px; line-height:2;}
.style4 {font-family: "標楷體"; line-height:2;}
.style5 {font-size: 18px; line-height:2;}
.style6 {font-family: "標楷體"; font-size: 18px; line-height:2; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
	line-height:2;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
	line-height:2;
}
.style9 {font-family: "標楷體"; line-height:2;}
.style10 {font-size: 16px; line-height:2;}
.style11 {font-size: 14px; line-height:2;}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
	line-height:2;
}
.style13 {font-size: 14px; font-family: "標楷體"; line-height:2; }
.style14 {
	font-size: 22px;
	font-family: "標楷體";
	line-height:1;
}
.style15 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style16 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style17 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style18 {font-family: "標楷體"; font-size: 20px; line-height:2; }
.style19 {font-size: 24px; line-height:2; }
.style20 {font-size: 36px; line-height:2; }
.style21 {font-size: 18px; line-height:2; }
.style22 {font-family: "標楷體"; font-size: 18px;}
.style23 {font-family: "標楷體"; font-size: 14px;}
.style24 {font-family: "標楷體"; font-size: 12px;}
.style25 {font-family: "標楷體"; font-size: 24px;}
.style26 {font-family: "標楷體"; font-size: 10px;}
.style27 {font-family: "@新細明體"; font-size: 10px;}
.style28 {font-family: "@新細明體"; font-size: 14px;}
.style29 {font-family: "@新細明體"; font-size: 14px;}
.style30 {font-family: "@新細明體"; font-size: 14px;}
.style31 {font-family: "@標楷體"; font-size: 14px;}
.pageprint {
  margin-left: 5mm;
  margin-right: 5.08mm;
  margin-top: 0mm;
  margin-bottom: 5.08mm;
}
-->
</style>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close
set rsUInfo=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

thenPasserUnit=""
strSQL="select UnitID,UnitTypeID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
	Sys_GroupUnitName=trim(rsunit("UnitName"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsunit=conn.Execute(strSQL)
Sys_UnitID=trim(rsunit("UnitID"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
else
	Sys_SendBillSN=request("hd_BillSN")
End if
BillSN=Split(Sys_SendBillSN,",")
BillState=""
for i=0 to Ubound(BillSN)
	For k=0 to 2
		If k=0 Then
			if trim(request("Sys_PasserJude"))="1" then '裁決書
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_chromat.asp"-->
					</Div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/PasserJudeDeliver_chromat.asp"-->
					</Div><%
				end if
			end if
		elseif k=1 then
			if trim(request("Sys_PasserUrge"))="1" then '催繳書
				if Sys_UnitID="0480" then
					if BillState<>"" then
						response.write "<div class=""PageNext"">&nbsp;</div>"%>
						<div id="L78" style="position:relative;">
						<!--#include virtual="traffic/PasserBase/PasserUrgeDeliver_chromat_6.asp"-->
						</div><%
					else
						BillState="1"%>
						<div id="L78" style="position:relative;">
						<!--#include virtual="traffic/PasserBase/PasserUrgeDeliver_chromat_6.asp"-->
						</div><%
					end if
				else
					if BillState<>"" then
						response.write "<div class=""PageNext"">&nbsp;</div>"%>
						<div id="L78" style="position:relative;">
						<!--#include virtual="traffic/PasserBase/PasserUrgeDeliver_chromat.asp"-->
						</div><%
					else
						BillState="1"%>
						<div id="L78" style="position:relative;">
						<!--#include virtual="traffic/PasserBase/PasserUrgeDeliver_chromat.asp"-->
						</div><%
					end if
				end if
			end if
		elseif k=2 then
			if trim(request("Sys_PasserDeliver"))="1" then '送達證書
				DeliverKind=1
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				end if
			elseif trim(request("Sys_PasserDeliver"))="2" then
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBase_Deliver_chromat.asp"-->
					</div><%
				end if
			elseif trim(request("Sys_PasserDeliver"))="3" then
				if BillState<>"" then
					response.write "<div class=""PageNext"">&nbsp;</div>"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBaseTaiChungCity_Deliver_chromat.asp"-->
					</div><%
				else
					BillState="1"%>
					<div id="L78" style="position:relative;">
					<!--#include virtual="traffic/PasserBase/BillBaseTaiChungCity_Deliver_chromat.asp"-->
					</div><%
				end if
			end if
		end if
	next
next

'===========================================================================================
Function wordporss(word)
	pro_word=""
	pro_wordtmp=""
	a=0
	for y=len(word) to 1 step -1

		If Asc(Mid(word,y,1))<0 Then 
			pro_word=pro_word&pro_wordtmp&Mid(word,y,1)
			if y<>1 then pro_word=pro_word&"<br>"
			a=0
			pro_wordtmp=""
		Else	
			If InStr(Mid(word,y,1),">")>0 Then a=0
			a=a+1
			if a=1 And Len(word)<>Y then pro_wordtmp="<br>"&pro_wordtmp  
			if a=1 And Len(word)=Y then pro_wordtmp=pro_wordtmp&"<br>"

			pro_wordtmp=Mid(word,y,1)&pro_wordtmp

		End If
		
	next
	wordporss=pro_word&pro_wordtmp
end Function


Function chstr(istr) ' 半形轉全形
Dim strtmp
strtmp = Replace(istr, "(", "（") 
strtmp = Replace(strtmp, ")", "）") 
strtmp = Replace(strtmp, "[", "〔") 
strtmp = Replace(strtmp, "]", "〕") 
strtmp = Replace(strtmp, "{", "｛") 
strtmp = Replace(strtmp, "}", "｝") 
strtmp = Replace(strtmp, ".", "。") 
strtmp = Replace(strtmp, ",", "，") 
strtmp = Replace(strtmp, ";", "；") 
strtmp = Replace(strtmp, ":", "：") 
strtmp = Replace(strtmp, "-", "－") 
strtmp = Replace(strtmp, "?", "？") 
strtmp = Replace(strtmp, "!", "！") 
strtmp = Replace(strtmp, "@", "＠") 
strtmp = Replace(strtmp, "#", "＃") 
strtmp = Replace(strtmp, "$", "＄") 
strtmp = Replace(strtmp, "%", "％") 
strtmp = Replace(strtmp, "&", "＆") 
strtmp = Replace(strtmp, "|", "｜") 
strtmp = Replace(strtmp, "", "＼") 
strtmp = Replace(strtmp, "/", "／") 
strtmp = Replace(strtmp, "+", "＋") 
strtmp = Replace(strtmp, "=", "＝") 
strtmp = Replace(strtmp, "*", "＊") 
strtmp = Replace(strtmp, "0", "０") 
strtmp = Replace(strtmp, "1", "１") 
strtmp = Replace(strtmp, "2", "２") 
strtmp = Replace(strtmp, "3", "３") 
strtmp = Replace(strtmp, "4", "４") 
strtmp = Replace(strtmp, "5", "５") 
strtmp = Replace(strtmp, "6", "６") 
strtmp = Replace(strtmp, "7", "７") 
strtmp = Replace(strtmp, "8", "８") 
strtmp = Replace(strtmp, "9", "９") 
strtmp = Replace(strtmp, "a", "ａ") 
strtmp = Replace(strtmp, "b", "ｂ") 
strtmp = Replace(strtmp, "c", "ｃ") 
strtmp = Replace(strtmp, "d", "ｄ") 
strtmp = Replace(strtmp, "e", "ｅ") 
strtmp = Replace(strtmp, "f", "ｆ") 
strtmp = Replace(strtmp, "g", "ｇ") 
strtmp = Replace(strtmp, "h", "ｈ") 
strtmp = Replace(strtmp, "i", "ｉ") 
strtmp = Replace(strtmp, "j", "ｊ") 
strtmp = Replace(strtmp, "k", "ｋ") 
strtmp = Replace(strtmp, "l", "ｌ") 
strtmp = Replace(strtmp, "m", "ｍ") 
strtmp = Replace(strtmp, "n", "ｎ") 
strtmp = Replace(strtmp, "o", "ｏ") 
strtmp = Replace(strtmp, "p", "ｐ") 
strtmp = Replace(strtmp, "q", "ｑ") 
strtmp = Replace(strtmp, "r", "ｒ") 
strtmp = Replace(strtmp, "s", "ｓ") 
strtmp = Replace(strtmp, "t", "ｔ") 
strtmp = Replace(strtmp, "u", "ｕ") 
strtmp = Replace(strtmp, "v", "ｖ") 
strtmp = Replace(strtmp, "w", "ｗ") 
strtmp = Replace(strtmp, "x", "ｘ") 
strtmp = Replace(strtmp, "y", "ｙ") 
strtmp = Replace(strtmp, "z", "ｚ") 
strtmp = Replace(strtmp, "A", "Ａ") 
strtmp = Replace(strtmp, "B", "Ｂ") 
strtmp = Replace(strtmp, "C", "Ｃ") 
strtmp = Replace(strtmp, "D", "Ｄ") 
strtmp = Replace(strtmp, "E", "Ｅ") 
strtmp = Replace(strtmp, "F", "Ｆ") 
strtmp = Replace(strtmp, "G", "Ｇ") 
strtmp = Replace(strtmp, "H", "Ｈ") 
strtmp = Replace(strtmp, "I", "Ｉ") 
strtmp = Replace(strtmp, "J", "Ｊ") 
strtmp = Replace(strtmp, "K", "Ｋ") 
strtmp = Replace(strtmp, "L", "Ｌ") 
strtmp = Replace(strtmp, "M", "Ｍ") 
strtmp = Replace(strtmp, "N", "Ｎ") 
strtmp = Replace(strtmp, "O", "Ｏ") 
strtmp = Replace(strtmp, "P", "Ｐ") 
strtmp = Replace(strtmp, "Q", "Ｑ") 
strtmp = Replace(strtmp, "R", "Ｒ") 
strtmp = Replace(strtmp, "S", "Ｓ") 
strtmp = Replace(strtmp, "T", "Ｔ") 
strtmp = Replace(strtmp, "U", "Ｕ") 
strtmp = Replace(strtmp, "V", "Ｖ") 
strtmp = Replace(strtmp, "W", "Ｗ") 
strtmp = Replace(strtmp, "X", "Ｘ") 
strtmp = Replace(strtmp, "Y", "Ｙ") 
strtmp = Replace(strtmp, "Z", "Ｚ") 
strtmp = Replace(strtmp, " ", "　") 
chstr = strtmp 
End Function 
%>
<div id=idDiv></div>
<div id=oCodeDiv></div>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	for(i=0;i<document.all("idDiv").length-1;i++){
		document.all("idDiv")[i].style.filter="progid:DXImageTransform.Microsoft.BasicImage(rotation=3)";
	}
	window.focus();
	printWindow(true,4.23,4.23,4.23,4.23);
</script>