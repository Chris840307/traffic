<!--#include virtual="traffic/takecar/Common/DB.ini"-->
<!--#include virtual="traffic/takecar/Common/AllFunction.inc"-->
<%
	fname="拖吊件數_"&CDbl(year((now))-1911)&Right("0"&Month((date)),2)&Right("0"&Day((date)),2)&"包含隆田保管場.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	Response.contenttype="application/x-msexcel; charset=MS950" 

SDate = gOutDT(request("tDate"))
EDate = gOutDT(request("tDate2"))

%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>拖吊件數</title>
</head>

<body>
<P>&nbsp;</p>

<table border="1" cellspacing="0" cellpadding="0" width="100%">
		<td align="center" colspan="9"><b><font face="標楷體" style="font-size:16pt;">臺南市政府警察局交通警察大隊<%=year(SDate)-1911%>年<%=right("0"&month(SDate),2)%>月<%=right("0"&Day(SDate),2)%>日至<%=right("0"&month(EDate),2)%>月<%=right("0"&Day(EDate),2)%>日<br>拖吊件數</font></b></td>
<tr>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">月份</font></b></td>
		<td align="center" style="width:80px"><b><font face="標楷體" style="font-size:12pt;">類別</b></td>
		<td align="center" style="width:50px"><b><font face="標楷體" style="font-size:12pt;">車輛</b></td>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">溪南<BR>拖吊場</b></td>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">永康<BR>拖吊場</b></td>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">新營<BR>拖吊場</b></td>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">隆田<BR>保管場</b></td>
		<td align="center" style="width:70px"><b><font face="標楷體" style="font-size:12pt;">合計</b></td>
		<td align="center" style="width:100px"><b><font face="標楷體" style="font-size:12pt;">總計</b></td>
<tr>
<%
function Getcnt(InCarTypeID,CarTypeID,UnitID,SDate,eDate)
tmpcnt=0


  strsql="Select count(sn) as cnt from Takebase where RecordStateid='0' and InCarTypeID in ("&InCarTypeID&") and CarTypeID in ("&CarTypeID&") and NowKeepUnitID in ("&UnitID&") and indatetime between to_date('"&SDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and to_date('"&eDate&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') "
'response.write "typeid"&typeid&"<br>"&strsql
  set rstmp=conn.execute(strsql)
  if not rstmp.eof then tmpcnt=rstmp("cnt")
  set rstmp=nothing
Getcnt=tmpcnt
end function
AIC1=0 : AIB1=0
AOC1=0 : AOB1=0

AIC2=0 : AIB2=0
AOC2=0 : AOB2=0

AIC3=0 : AIB3=0
AOC3=0 : AOB3=0

AICT=0 : AIBT=0
AOCTT=0 : AOBT=0
AICBT=0 : AOCBT=0

%>
<%for I=cdbl(month(SDate)) to cdbl(month(EDate))
IC1=0 : IB1=0
OC1=0 : OB1=0

IC2=0 : IB2=0
OC2=0 : OB2=0

IC3=0 : IB3=0
OC3=0 : OB3=0

ICT=0 : IBT=0
OCTT=0 : OBT=0
ICBT=0 : OCBT=0
If i=Month(SDate) And i=Month(EDate) Then
	SDate2 = Year(SDate)&"/"&month(SDate)&"/"&day(SDate)
	EDate2 = Year(EDate)&"/"&month(EDate)&"/"&day(EDate)
elseIf i=Month(SDate) Then 

		SDate2 = Year(SDate)&"/"&month(SDate)&"/01"
		NewDate = DateAdd("m", 1, SDate2)
		NewDate = DateAdd("d", -1, NewDate)
		EDate2 = NewDate
		SDate2 = Year(SDate)&"/"&month(SDate)&"/"&day(SDate)
ElseIf i=Month(EDate) Then 
	    SDate2 = Year(EDate)&"/"&month(EDate)&"/01"
		eDate2=eDate
Else
		SDate2 = year(SDate)&"/"&i&"/01"
		NewDate = DateAdd("m", 1, SDate2)
		NewDate = DateAdd("d", -1, NewDate)
		EDate2 = NewDate
End If

%>
<td rowspan="4" align="center"><font face="標楷體" style="font-size:12pt;"><%=I%>月份</td>
<td rowspan="2" align="center"><font face="標楷體" style="font-size:12pt;">違規停車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;">汽車</td>
	<%
IC1=cdbl(Getcnt("'1'","'A','D','E','F','G','H','I','J','K'","'07C1'",SDate2,EDate2))
AIC1=AIC1+IC1
IC2=cdbl(Getcnt("'1'","'A','D','E','F','G','H','I','J','K'","'07D4'",SDate2,EDate2))
AIC2=AIC2+IC2
IC3=cdbl(Getcnt("'1'","'A','D','E','F','G','H','I','J','K'","'07D3'",SDate2,EDate2))
AIC3=AIC3+IC3
IC4=cdbl(Getcnt("'1'","'A','D','E','F','G','H','I','J','K'","'07D1','07D5'",SDate2,EDate2))
AIC4=AIC4+IC4
ICT=cdbl(IC1)+cdbl(IC2)+cdbl(IC3)+cdbl(IC4)
AICT=AICT+ICT

IB1=cdbl(Getcnt("'1'","'B','C'","'07C1'",SDate2,EDate2))
AIB1=AIB1+IB1
IB2=cdbl(Getcnt("'1'","'B','C'","'07D3'",SDate2,EDate2))
AIB2=AIB2+IB2
IB3=cdbl(Getcnt("'1'","'B','C'","'07D4'",SDate2,EDate2))
AIB3=AIB3+IB3
IB4=cdbl(Getcnt("'1'","'B','C'","'07D1','07D5'",SDate2,EDate2))
AIB4=AIB4+IB4
IBT=cdbl(IB1)+cdbl(IB2)+cdbl(IB3)+cdbl(IB4)
AIBT=AIBT+IBT
ICBT=cdbl(ICT)+cdbl(IBT)
AICBT=AICBT+ICBT

OC1=cdbl(Getcnt("'2','3','4','5','6'","'A','D','E','F','G','H','I','J','K'","'07C1'",SDate2,EDate2))
AOC1=AOC1+OC1
OC2=cdbl(Getcnt("'2','3','4','5','6'","'A','D','E','F','G','H','I','J','K'","'07D4'",SDate2,EDate2))
AOC2=AOC2+OC2
OC3=cdbl(Getcnt("'2','3','4','5','6'","'A','D','E','F','G','H','I','J','K'","'07D3'",SDate2,EDate2))
AOC3=AOC3+OC3
OC4=cdbl(Getcnt("'2','3','4','5','6'","'A','D','E','F','G','H','I','J','K'","'07D1','07D5'",SDate2,EDate2))
AOC4=AOC4+OC4
OCTT=cdbl(OC1)+cdbl(OC2)+cdbl(OC3)+cdbl(OC4)
AOCT=AOCT+OCTT

OB1=cdbl(Getcnt("'2','3','4','5','6'","'B','C'","'07C1'",SDate2,EDate2))
AOB1=AOB1+OB1
OB2=cdbl(Getcnt("'2','3','4','5','6'","'B','C'","'07D4'",SDate2,EDate2))
AOB2=AOB2+OB2
OB3=cdbl(Getcnt("'2','3','4','5','6'","'B','C'","'07D3'",SDate2,EDate2))
AOB3=AOB3+OB3
OB4=cdbl(Getcnt("'2','3','4','5','6'","'B','C'","'07D1','07D5'",SDate2,EDate2))
AOB4=AOB4+OB4
OBT=cdbl(OB1)+cdbl(OB2)+cdbl(OB3)+cdbl(OB4)
AOBT=AOBT+OBT
OCBT=cdbl(OCTT)+cdbl(OBT)
AOCBT=AOCBT+OCBT
%>

    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IC1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IC2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IC3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IC4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=ICt%></td>
    <td align="center" rowspan="2"><font face="標楷體" style="font-size:12pt;"><%=ICBt%></td>
   <tr>
    <td align="center" ><font face="標楷體" style="font-size:12pt;">機車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IB4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=IBT%></td>
   <tr>
<td rowspan="2" align="center"><font face="標楷體" style="font-size:12pt;">其他</td>
<td align="center" ><font face="標楷體" style="font-size:12pt;">汽車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OC1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OC2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OC3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OC4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OCTT%></td>
    <td align="center" rowspan="2"><font face="標楷體" style="font-size:12pt;"><%=OCBt%></td>
   <tr>
<td align="center" ><font face="標楷體" style="font-size:12pt;">機車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OB4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=OBT%></td>
   <tr>

<%next%>
<td rowspan="4" align="center"><font face="標楷體" style="font-size:12pt;">合計</td>
<td rowspan="2" align="center"><font face="標楷體" style="font-size:12pt;">違規停車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;">汽車</td>

    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AICt%></td>
    <td align="center" rowspan="2"><font face="標楷體" style="font-size:12pt;"><%=AICBt%></td>
   <tr>
    <td align="center" ><font face="標楷體" style="font-size:12pt;">機車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIBT%></td>
   <tr>
<td rowspan="2" align="center"><font face="標楷體" style="font-size:12pt;">其他</td>
<td align="center" ><font face="標楷體" style="font-size:12pt;">汽車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOC1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOC2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOC3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOC4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOCT%></td>
    <td align="center" rowspan="2"><font face="標楷體" style="font-size:12pt;"><%=AOCBt%></td>
   <tr>
<td align="center" ><font face="標楷體" style="font-size:12pt;">機車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOB4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AOBT%></td>
   <tr>
<td rowspan="3" align="center" style="Height:100px;" valign="center"><font face="標楷體" style="font-size:12pt;">違規停<br>車及其<br>他拖吊<br>總合</td>
    <td align="center" colspan="2"><font face="標楷體" style="font-size:12pt;">汽車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC1+AOC1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC2+AOC2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC3+AOC3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC4+AOC4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AICt+AOCT%></td>
    <td align="center" rowspan="2"><font face="標楷體" style="font-size:12pt;"><%=AICBt+AOCBt%></td>
   <tr>
    <td align="center" colspan="2" style="Height:40px"><font face="標楷體" style="font-size:12pt;">機車</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB1+AOB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB2+AOB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB3+AOB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIB4+AOB4%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIBT+AOBT%></td>
<tr>
    <td align="center" colspan="2"><font face="標楷體" style="font-size:12pt;">總計</td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC1+AOC1+AIB1+AOB1%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC2+AOC2+AIB2+AOB2%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC3+AOC3+AIB3+AOB3%></td>
    <td align="center" ><font face="標楷體" style="font-size:12pt;"><%=AIC4+AOC4+AIB4+AOB4%></td>
    <td align="center"><font face="標楷體" style="font-size:12pt;"><%=AICBt+AOCBt%></td>
    <td align="center"><font face="標楷體" style="font-size:12pt;">&nbsp;</td>
</table>
</center>

</body>
<%conn.close%>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
//	window.focus();
	//printWindow3(true,5,5.08,1,5.08);
window.close();
</script>

</html>
