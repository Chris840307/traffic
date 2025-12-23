<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_戶籍地址補正資料列表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

Server.ScriptTimeout = 68000
Response.flush

'權限
'AuthorityCheck(234)
RecordDate=split(gInitDT(date),"-")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="CarDataSelect" then
	strwhere=Session("PrintCarDataSQL")&" and a.CarNo like '%"&trim(request("SelCarNo"))&"%'"
else
	strwhere=Session("PrintCarDataSQL")	
end if
	'Session.Contents.Remove("PrintCarDataSQLxls")
	'Session("PrintCarDataSQLxls")=strwhere	
	dcitype=trim(request("dcitype"))

	strdata=" and (substr(e.ownerid,2,1)<>'A' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'S' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'D' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'F' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'G' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'H' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'J' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'K' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'L' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'Z' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'X' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'C' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'V' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'B' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'N' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'M' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'Q' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'W' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'E' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'R' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'T' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'Y' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'U' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'I' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'O' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>'P' "
	strdata=strdata&" and substr(e.ownerid,2,1)<>' ' "
	strdata=strdata&")"

	strdata2=" and (substr(e.ownerid,1,1)='A' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='S' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='D' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='F' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='G' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='H' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='J' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='K' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='L' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='Z' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='X' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='C' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='V' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='B' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='N' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='M' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='Q' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='W' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='E' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='R' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='T' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='Y' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='U' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='I' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='O' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)='P' "
	strdata2=strdata2&" or substr(e.ownerid,1,1)=' ' "
	strdata2=strdata2&")"

	strSQL="select distinct e.billno,e.ownerid,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='') "&strwhere&" order by a.RecordDate"

	set rsfound=conn.execute(strSQL)

	strCnt="select count(*) as cnt from (select distinct a.SN,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='')"&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere
%>
單號	車號	證號	姓名	戶籍地址
<%
If Not rsfound.Bof Then rsfound.MoveFirst 
	While Not rsfound.Eof
		Response.Write rsfound("BillNo")&"	"&rsfound("CarNo")&"	"&rsfound("OwnerID")&"	"&funcCheckFont(rsfound("Owner"),20,1)&vbnewline
		Response.flush
		rsfound.MoveNext
	Wend
	rsfound.close
	set rsfound=nothing
conn.close%>