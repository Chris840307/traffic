<!--#include virtual="traffic/takecar/Common/css.txt"-->
<!--#include virtual="traffic/takecar/Common/DB.ini"-->
<!--#include virtual="traffic/takecar/Common/AllFunction.inc"-->
<!--#include virtual="traffic/takecar/Common/Login_Check.asp"--> 
<%

Function GetCarTypeName(CarTypeID)
	tmp=""
				sql="select CarTypeName from MoveCost where CarTypeid='"&CarTypeID&"'"
				Set rst=conn.execute(sql)
				If Not rst.eof Then tmp=rst("CarTypeName")
GetCarTypeName=tmp
End function

'fname=GetTakeCarUnitName(Session("Unit_ID"))&"_"&GetCdate(now)&"公告清冊.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'Response.contenttype="application/x-msexcel; charset=MS950" 

	
	SDate=gOutDT(request("tDate"))
	EDate=gOutDT(request("tDate2"))

%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>未領車輛清單</title>
</head>

<body>

<p align="center"><b><font face="標楷體" style="font-size:20pt;">臺南市政府警察局交通警察大隊拖吊場逾期未領車輛公告清冊</font></b>
<center>
<table border="0" width="100%">
<td align="center" colspan="10"><font face="標楷體" style="font-size:12pt;">[<%=GetTakeCarUnitName(request("Sys_UnitID"))%>]<br>統計日期：
	<%If SDate = EDate then
	  response.write GetCDate(SDate)
	else
	  response.write GetCDate(SDate)&"～"&GetCDate(EDate)
	End if%>
</font>
</td>
<tr>

<td align="left" colspan="10">
<font face="標楷體" style="font-size:12pt;">列印日期：<%=GetCdate(now)%>&nbsp;&nbsp;&nbsp;&nbsp;第 1 頁
</td>
</table>
<%
PageCnt=1

tmpFontSize="12"
tmpLineHeight="16"
%>

<table border="1" cellspacing="0" cellpadding="0" width="100%">

		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">車種</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">車牌號碼</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">引擎號碼</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">廠牌</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">顏色</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">進場時間</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">拖吊路段</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">備考</font></td>
	</tr>  
	<%

	
	where=" between to_date('"&sDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and to_date('"&EDate&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and a.recordstateid='0' and TakeStatus<>'4' "

	where=where&" and a.sn in ("&request("sn")&") and inCarTypeID<>'5' "
	sql="select * from takebase a where indatetime"&where&" ORDER BY INDATETIME"
'response.write sql
    Set rs=conn.execute(sql)
	sn=0

		while Not rs.eof
			sn=sn+1
			If sn Mod 25=0 Then '數值
			PageCnt=PageCnt+1
			response.write "</table>"
			response.write "<div class=""PageNext"">&nbsp;</div>"
			%>
				<p align="center"><b><font face="標楷體" style="font-size:20pt;">臺南市政府警察局交通警察大隊拖吊場逾期未領車輛公告清冊</font></b>
				<center>
				<table border="0" width="100%">
				<td align="center" colspan="10"><font face="標楷體" style="font-size:12pt;">[<%=GetTakeCarUnitName(request("Sys_UnitID"))%>]<br>統計日期：
					<%If SDate = EDate then
					  response.write GetCDate(SDate)
					else
					  response.write GetCDate(SDate)&"～"&GetCDate(EDate)
					End if%>
				</font>
				</td>
				<tr>

				<td align="left" colspan="10">
				<font face="標楷體" style="font-size:12pt;">列印日期：<%=GetCdate(now)%>&nbsp;&nbsp;&nbsp;&nbsp;第 <%=PageCnt%> 頁
				</td>
				</table>


				<table border="1" cellspacing="0" cellpadding="0" width="100%">

						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">車種</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">車牌號碼</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">引擎號碼</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">廠牌</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">顏色</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">進場時間</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">拖吊路段</font></td>
						<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">備考</font></td>
					</tr>  
			<%
			End if
			response.write "<tr>"
           ' response.write "<td align='center'style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt"">&nbsp;"&sn&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt"" NOWRAP><font face=""標楷體"">&nbsp;"&GetCarTypeToName(rs("CarTypeID"))&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt"" NOWRAP><font face=""標楷體"">&nbsp;"&rs("CarNo")&"</td>"            
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt""  NOWRAP><font face=""標楷體"">&nbsp;"&rs("EngineNum")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt""  NOWRAP><font face=""標楷體"">&nbsp;"&rs("Brand")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt""  NOWRAP><font face=""標楷體"">&nbsp;"&rs("color")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt"" NOWRAP><font face=""標楷體"">&nbsp;"&GetCdatetime(rs("InDateTime"))&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt""><font face=""標楷體"">&nbsp;"&rs("Area")&rs("TakePlace")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;Line-height:"&tmpLineHeight&"pt""><font face=""標楷體"">&nbsp;"&rs("Notes")&"</td>"
			rs.movenext
		wend
    Set rs=Nothing
	
	sql="select distinct a.CarTypeID from takebase a where a.InDateTime "&where&" order by a.CarTypeID"
'response.write sql
CarTypeID=""
		Set rs=conn.execute(sql)
		while Not rs.eof 
		    CarTypeID=CarTypeID&","&rs("CarTypeID")
		  rs.MoveNext
		wend

'----------------車種---------------------------------------------------------------------------------------------------------------------


CarTypeID=Split(CarTypeID,",")

	%>
	<tr>
	<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">統計事項</font></td>
<td align="center" style="border-bottom-style : solid;"  colspan="11"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;Line-height:<%=tmpLineHeight%>pt">
&nbsp;
<%
tmpcnt=0
totaltmpcnt=0
for i=1 to ubound(CarTypeID)

sql="select count(a.sn) as cnt from takebase a where a.InDateTime "
sql=sql&where&" and a.CarTypeID ='"&CarTypeID(i)&"' and a.sn in ("&request("sn")&")"

Set rstmp2=conn.execute(sql)

if not rstmp2.eof then tmpcnt=cdbl("0"&rstmp2("cnt"))
totaltmpcnt=totaltmpcnt+tmpcnt

response.write "&nbsp;"&GetCarTypeName(cartypeid(i))&":"&tmpcnt
next
response.write "&nbsp;合計:"&totaltmpcnt
set rstmp2=nothing%>
</td>
</table>
</body>
<%

conn.close

%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow3(true,8,5.08,1,5.08);
</script>

</html>
