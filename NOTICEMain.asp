<!-- #include file="Common\db.ini" -->
<!-- #include file="Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">

<title>
公告訊息
</title>
</head>
<%

set fs=Server.CreateObject("Scripting.FileSystemObject")

tdate ="select to_char(sysdate,'yymmdd') as tdate from dual"
set rstdate =conn.execute(tdate )
tdate =trim(rstdate ("tdate"))
rstdate.close

	ArgueDate1=DateAdd("d",-10,date) & " 0:0:0"
	ArgueDate2=date & " 23:29:59" 

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))

rsCity.close

sStartDate=gOutDT(ginitdt(now)) & " 00:00:00 "
sEndDate=gOutDT(ginitdt(now)) & " 23:59:59 "
strSQL="Select count(*) billcount from BillBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and RecordStateID <> -1 "

set rssysinfo=conn.execute(strSQL)
billcount=rssysinfo("billcount") 
set rssysinfo=nothing

strSQL="select count(ExchangeTypeID) as Num from DCILog Where ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID")  & " and ExchangeTypeID='W' "
set rssysinfo=conn.execute(strSQL)
Sendcount=rssysinfo("Num") 
set rssysinfo=nothing

strwhere=" and a.RecordMemberID="& Session("User_ID") & " and a.ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS') and a.ExchangeTypeID='W'"
if sys_City="基隆市" then
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') or usetool=8) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool=8 "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if


		set chksuess=conn.execute(strSQL)
		filsuess=CDbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='-1' "&strwhere

		set chksuess=conn.execute(strSQL)
		fildel=CDbl(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=CDbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=CDbl(Dbrs("cnt"))
		Dbrs.close

		if sys_City="基隆市" then
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt=CDbl(Dbrs("cnt"))
		Dbrs.close


'---------------------------------------------------------------------------------------------------------------


strSQL="select count(ExchangeTypeID) as Num from DCILog Where ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and ExchangeTypeID='W' "
set rssysinfo=conn.execute(strSQL)
Sendcount2=rssysinfo("Num") 
set rssysinfo=nothing

strSQL="select ChName from MemberData Where MemberID=" &  Session("User_ID") 
set rssysinfo=conn.execute(strSQL)
UserName=rssysinfo("ChName") 
set rssysinfo=nothing

strwhere="and a.RecordMemberID="& Session("User_ID") & " and a.ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS') and a.ExchangeTypeID='W'"
if sys_City="基隆市" then
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') or usetool=8) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool=8 "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if


		set chksuess=conn.execute(strSQL)
		filsuess2=CDbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='-1' "&strwhere

		set chksuess=conn.execute(strSQL)
		fildel2=CDbl(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum2=CDbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata2=CDbl(Dbrs("cnt"))
		Dbrs.close

		if sys_City="基隆市" then
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt2=CDbl(Dbrs("cnt"))
		Dbrs.close
set Dbrs=nothing

%>

<table border="0" width="100%" id="table1">
	<tr bgcolor="#FFCC33">
		<td>公告訊息</td>
	</tr>
	
		
<%
FileName=Server.MapPath(fs.GetFileName("note"& tdate &".txt"))

	    if fs.fileExists(FileName)=true then
           set txtStream = fs.opentextfile(FileName) 
              txtline = txtStream.readAll
              response.write "<tr><td><font color=red size=""4"">"&txtline&"</font></td></tr>"
     	end if

 set txtStream = nothing
 set fs = nothing 
%>
		
	
</table>
<%	

	strDelErr="select * from Dcilog where ExchangeTypeID='E' and (DciReturnStatusID<>'S' or DciReturnStatusID is null)" &_
		" and ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&Session("User_ID")
	set rsDelErr=conn.execute(strDelErr)
	if not rsDelErr.eof then
%>
	<table border="1" width="100%" id="table3">
		<tr bgcolor="#FFCC33">
			<td colspan="4">十日內刪除未處理、異常</td>
		</tr>
		<tr bgcolor="#FFCC33">
			<td width="20%">上傳日期</td>
			<td width="20%">批號</td>
			<td width="20%">單號</td>
			<td width="40%">訊息</td>
		</tr>
<%
	end if
	If Not rsDelErr.Bof Then rsDelErr.MoveFirst 
	While Not rsDelErr.Eof
%>
		<tr>
			<td>
			<%=year(rsDelErr("ExchangeDate"))-1911&right("00"&month(rsDelErr("ExchangeDate")),2)&right("00"&day(rsDelErr("ExchangeDate")),2)%>
			</td>
			<td>
			<%=rsDelErr("Batchnumber")%>
			</td>
			<td>
			<%=rsDelErr("BillNo")%>
			</td>
			<td>
			<%
			if trim(rsDelErr("DciReturnStatusID"))="" or isnull(rsDelErr("DciReturnStatusID")) then
				response.write "未處理"
			else
				strErr="select StatusContent from DciReturnStatus where DciActionID='E'" &_
					" and DciReturn='"&trim(rsDelErr("DciReturnStatusID"))&"' " 
				set rsErr=conn.execute(strErr)
				if not rsErr.eof then
					response.write rsErr("StatusContent")
				end if
				rsErr.close
				set rsErr=nothing
			end if
			%>
			</td>
		</tr>
<%
	rsDelErr.MoveNext
	Wend
	if not rsDelErr.eof then
%>
	</table>
<%
	end if
	rsDelErr.close
	set rsDelErr=nothing
%>
	 

<iframe name="I1" src="NoiceMain_Data.asp" width="100%" height="257" frameBorder="0">
您的瀏覽器不支援內置框架或目前的設定為不顯示內置框架。</iframe></p>
<%
	if sys_City="台中市" or sys_City="台中縣" then

		sStartDate=gOutDT(ginitdt(now-1)) & " 00:00:00 "
		sEndDate=gOutDT(ginitdt(now-1)) & " 23:59:59 "
		strSQL="Select count(*) billcount2 from BillBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and RecordStateID <> -1 "

		set rssysinfo=conn.execute(strSQL)
		billcount2=rssysinfo("billcount2") 
		set rssysinfo=nothing	
%>
<table border="0" width="100%" id="table2" bgcolor="#FFCC33">
	<tr>
		<td><%=UserName%>&nbsp; 處理進度 (1~68條)</td>
	</tr>
</table>
<br>
昨日建檔：<%=billcount2%>&nbsp;&nbsp;入案：<%=Sendcount2%>&nbsp;&nbsp;成功：<%=filsuess2%>&nbsp;&nbsp;異常：<%=fildel2+errCatCnt2%>&nbsp;&nbsp;未處理：<%=DBsum2-CDbl(filsuess2)-CDbl(fildel2)-CDbl(deldata2)-CDbl(errCatCnt2)%><br>
今日建檔：<%=billcount%>&nbsp;&nbsp;入案：<%=Sendcount%>&nbsp;&nbsp;成功：<%=filsuess%>&nbsp;&nbsp;異常：<%=fildel+errCatCnt%>&nbsp;&nbsp;未處理：<%=DBsum-CDbl(filsuess)-CDbl(fildel)-CDbl(deldata)-CDbl(errCatCnt)%><br>
<%
	end if
%>

<%
if sys_City="花蓮縣" then
	strUlErr="select distinct(Batchnumber) from Dcilog a,DciReturnStatus b " &_
		" where (b.DCIreturnStatus=-1 or a.DciReturnStatusID is null)" &_
		" and a.ExchangeTypeID=b.DCIActionID(+) and a.DCIReturnStatusID=b.DCIReturn(+)" &_
		" and a.ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and a.RecordMemberID="&Session("User_ID")
		'response.write strUlErr
	set rsUErr=conn.execute(strUlErr)
	if not rsUErr.eof then
%>
	<table border="1" width="100%" id="table3">
		<tr bgcolor="#FF3300">
			<td colspan="4">十日內上傳未處理、異常批號</td>
		</tr>
<%
	end if
	If Not rsUErr.Bof Then rsUErr.MoveFirst 
	While Not rsUErr.Eof
%>
		<tr>
			<td>
			<%=rsUErr("Batchnumber")%>
			</td>
		</tr>
<%
	rsUErr.MoveNext
	Wend
	if not rsUErr.eof then
%>

	</table>

<%
	end if
	rsUErr.close
	set rsUErr=nothing
elseif sys_City="金門縣" and trim(Session("Group_ID"))="200" then
	strUlErr="select distinct(Batchnumber) from Dcilog a,DciReturnStatus b " &_
		" where (b.DCIreturnStatus=-1 or a.DciReturnStatusID is null)" &_
		" and a.ExchangeTypeID=b.DCIActionID(+) and a.DCIReturnStatusID=b.DCIReturn(+)" &_
		" and a.ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		'response.write strUlErr
	set rsUErr=conn.execute(strUlErr)
	if not rsUErr.eof then
%>
	<table border="1" width="100%" id="table3">
		<tr bgcolor="#FF3300">
			<td colspan="4">十日內上傳未處理、異常批號</td>
		</tr>
<%
	end if
	If Not rsUErr.Bof Then rsUErr.MoveFirst 
	While Not rsUErr.Eof
%>
		<tr>
			<td>
			<%=rsUErr("Batchnumber")%>
			</td>
		</tr>
<%
	rsUErr.MoveNext
	Wend
	if not rsUErr.eof then
%>

	</table>

<%
	end if
	rsUErr.close
	set rsUErr=nothing

end if
%>
<br>
<table border="0" width="100%" id="table4" >
	<tr>
		<td>
			
					<%
						
							strSQL="select a.UnitID,b.UnitName,a.ChName,a.cnt from (select UnitID,ChName,count(*) cnt from MemberData where  ACCOUNTSTATEID=0 and RECORDSTATEID=0 group by UnitID,ChName) a,UnitInfo b where a.cnt>1 and a.UnitID=b.UnitID order by UnitName"
							set rs=conn.execute(strSQL)
							if Not rs.eof then 
								Response.Write "<font size='3'><b> 下列人員於同一單位內姓名重覆 ， 請 各單位承辦 確認使用中帳號，<br>' 停用 ' 非使用中的帳號，避免資料有誤 .</font></br></br>"
								while Not rs.eof 
									if rs("ChName") <> "公積金" and rs("ChName") <> "陳俊良" then
										response.write rs("ChName") & "<img src='space.gif' width='15' height='8'>"  &  rs("UnitName") & "<img src='space.gif' width='35' height='8'>" & rs("cnt")& "<img src='space.gif' width='15' height='8'>" &  " 筆同時啟用中的人員資料 <br>" 
									end if
									rs.movenext
								wend
							end if
							rs.close
							
							
							
					%></td></tr>
</table>