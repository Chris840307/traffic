<!--#include virtual="traffic/takecar/Common/css.txt"-->
<!--#include virtual="traffic/takecar/Common/DB.ini"-->
<!--#include virtual="traffic/takecar/Common/AllFunction.inc"-->
<!--#include virtual="traffic/takecar/Common/Login_Check.asp"--> 
<%



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
<title>公告清冊</title>
</head>

<body>
<%tmpFontSize="8"%>
<p align="center">
<br><center></center>
<table border="1" cellspacing="0" cellpadding="0" width="100%">
		<td colspan="12" align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:16pt;line-height:28px;">臺南市政府警察局<%=year(EDate)-1911&"年"&Right("00"&month(EDate),2)%>月份交通違規案件移置保管逾期未領車輛公告清冊</font></td>
	<tr>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">編號</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">通知單<br>編號</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">車種</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">車牌號碼</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">引擎號碼</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">廠牌</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">車主姓名</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">車主地址</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">告發單<br>編號</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">進場<br>日期</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">違規<br>事項</font></td>
		<td align="center" style="border-bottom-style : solid;"><font face="標楷體" style="font-size:<%=tmpFontSize%>pt;">保管處所</font></td>
	</tr>  
	<%

	
	where=" between to_date('"&SDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and to_date('"&EDate&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and a.recordstateid='0' and TakeStatus<>'4' "

	where=where&" and a.sn in ("&request("sn")&")"
	sql="select * from takebase a where indatetime"&where&" order by indatetime"
'response.write sql
    Set rs=conn.execute(sql)
	sn=0

		while Not rs.eof
			sn=sn+1
			response.write "<tr>"
            response.write "<td align='center'style=""font-size:"&tmpFontSize&"pt;line-height:16px;"">&nbsp;"&sn&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("KeepNo")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&GetCarTypeToName(rs("CarTypeID"))&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("CarNo")&"</td>"            

            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("EngineNum")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("Brand")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("ownerName")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("owneraddr")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&rs("Billno")&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&GetCdate(rs("InDateTime"))&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&GetInCarTypeNameR(rs("IncarTypeID"),rs("RuleCode"))&"</td>"
            response.write "<td style=""font-size:"&tmpFontSize&"pt;"">&nbsp;"&GetTakeCarUnitName(rs("NowKeepUnitID"))&"</td>"
			rs.movenext
		wend
    Set rs=Nothing
	
	%>

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
	printWindow3(true,5,5.08,1,5.08);
</script>

</html>
