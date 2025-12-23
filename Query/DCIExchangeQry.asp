<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI 資料交換紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.btn3{
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</head>
<%
'檢查是否可進入本系統

If isEmpty(request("DB_Display")) Then
	Sys_Now=DateAdd("d",-2,date)&" "&hour(time)&":"&Minute(time)&":"&Second(time)

	Sys_Now2=DateAdd("d",-10,date)&" "&hour(time)&":"&Minute(time)&":"&Second(time)
	strSQL="select distinct a.batchnumber from DCILog a,DCIReturnStatus b where a.ExchangeTypeID=b.DCIActionID(+) and a.DCIReturnStatusID=b.DCIReturn(+) and b.DCIReturnStatus is null and a.ExchangeDate between TO_DATE('"&Sys_Now2&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&Sys_Now1&"','YYYY/MM/DD/HH24/MI/SS') and substr(a.batchnumber,1,1)<>'A' and a.RecordMemberID ="&Session("User_ID")

	chkbat=""

	set rschk=conn.execute(strSQL)
	while not rschk.eof
		If Not ifnull(chkbat) then chkbat=chkbat&"\n"
		chkbat=chkbat&rschk("batchnumber")
		rschk.movenext
	wend
	rschk.close
	If not ifnull(chkbat) Then
		Response.write "<script>"
		Response.Write "alert('您下列批號尚未回傳，請盡速確認！\n"&chkbat&"');"
		Response.write "</script>"
	End if
End if

Dim RecordDate,RecordDate1,strwhere,tmp_BatchNumber,Sys_BatchNumber,DB_Display

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

RecordDate=split(gInitDT(date),"-")
strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
if UCase(trim(request("Sys_BatchNumber")))<>"" then
	tmp_BatchNumber=split(UCase(trim(request("Sys_BatchNumber"))),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
	next
	'strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
end if
if request("DB_state")="Del" then
	if trim(request("SN"))="delall" then
		sql1 = "Delete From DCILog Where DCIReturnStatusID is null and BatchNumber in('"&Sys_BatchNumber&"')"
		ConnExecute sql1,0
		Conn.Execute(sql1)
		Response.write "<script>"
		Response.Write "alert('刪除完成！');"
		Response.write "</script>"
	else
		sql1 = "Delete From DCILog Where SN="&request("SN")&" and DCIReturnStatusID is null"
		ConnExecute sql1,0
		Conn.Execute(sql1)
		Response.write "<script>"
		Response.Write "alert('刪除完成！');"
		Response.write "</script>"
	end if
end If 

if request("DB_state")="CloseToSuss" then

	sql1 = "update DCILog set DCIReturnStatusID='S' Where SN="&request("SN")&" and exchangetypeid='N'"

	Conn.Execute(sql1)

	sql1 = "update BillBaseDcireturn set Status='S' Where BillNo=(select BillNo from Dcilog where SN="&request("SN")&") and carno=(select carno from dcilog where sn="&request("SN")&") and exchangetypeid='N'"

	Conn.Execute(sql1)

	
	If instr(",3,4,5,7,",trim(request("Hidden_ReturnMarkType"))) >0 Then

		sql1 = "update billbase set billstatus='"&Request("Hidden_ReturnMarkType")&"' Where sn=(select billsn from Dcilog where SN="&request("SN")&") and recordstateid=0"

		Conn.Execute(sql1)	
	elseIf instr(",Y,",trim(request("Hidden_ReturnMarkType"))) >0 Then

		sql1 = "update billbase set billstatus='3' Where sn=(select billsn from Dcilog where SN="&request("SN")&") and recordstateid=0"

		Conn.Execute(sql1)
		
	End if 

	ConnExecute "單退結案轉成功:BillSn="&Trim(request("SN")) ,355
	Response.write "<script>"
	Response.Write "alert('修改完成！');"
	Response.write "</script>"
end If 

if request("DB_state")="ReSend" then
	sql1 = "Update DCILog set FileName='',SeqNo='',dcireturnstatusid='',dcierrorcardata=null,dcierroriddata=null Where SN in (select a.SN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and (d.DCIReturnStatus='-1' or d.DCIReturnStatus is null) and a.BatchNumber in('"&Sys_BatchNumber&"'))"

	if instr(Sys_BatchNumber,"W")>0 then

		sql1=sql1&" and exists(select 'Y' from billbase where sn=dcilog.billsn and recordstateid=0)"
	end If 

	Conn.Execute(sql1)

	if instr(Sys_BatchNumber,"E")>0 then
		sql2 = "Update BillBase set BillStatus=6,RecordStateID=-1 Where SN in (select billsn from dcilog where BatchNumber in('"&Sys_BatchNumber&"'))"
		Conn.Execute(sql2)
	end if 

	if sys_City="苗栗縣" and instr(Sys_BatchNumber,"W")>0 then
		sql2 = "Update BillBase set CaseInDate2=sysdate Where SN in (select a.BillSn from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and (d.DCIReturnStatus='-1' or d.DCIReturnStatus is null) and a.BatchNumber in('"&Sys_BatchNumber&"'))"
		Conn.Execute(sql2)
	end if 

	Response.write "<script>"
	Response.Write "alert('重送完成！');"
	Response.write "</script>"
end if
if request("DB_Selt")="BatchSelt" then
	strwhere="":strDCISQL=""
	if UCase(request("Sys_BatchNumber"))<>"" then
		strDCISQL=" where BatchNumber in('"&Sys_BatchNumber&"')"
	end if

	if request("Sys_DCIReturnStatus_Batch")<>"" then
		strwhere=" and d.DCIReturnStatus "&request("Sys_DCIReturnStatus_Batch")
	end If 
	
	if request("Sys_CarType")="1" then
		strwhere=" and exists(select 'Y' from billbasedcireturn A1 where CarNo in(select CarNo from dcilog A2 where batchnumber in('"&Sys_BatchNumber&"')) and ExchangetypeID='A' and status='S' and ((mdrivername<>owner and mdrivername is not null) or (mdrivername is null and rdrivername is not null and rdrivername<>owner)) and carno=a.carno)"
	end if
	orderwhere=" order by a.Batchnumber,a.RecordDate"
end if
if request("DB_Selt")="Selt" then
	strwhere="":strDCISQL=""

	if request("RecordDate")<>"" and request("RecordDate1")<>""then
		strDCISQL=" where RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("ExchangeDate")<>"" and request("ExchangeDate1")<>""then
		ExchangeDate1=gOutDT(request("ExchangeDate"))&" 0:0:0"
		ExchangeDate2=gOutDT(request("ExchangeDate1"))&" 23:59:59"
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and ExchangeDate between TO_DATE('"&ExchangeDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ExchangeDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strDCISQL=" where ExchangeDate between TO_DATE('"&ExchangeDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ExchangeDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
	end if

	if trim(request("ExchangeDate_h"))<>"" or trim(request("ExchangeDate1_h"))<>"" then
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and to_char(ExchangeDate,'hh') between "&trim(request("ExchangeDate_h"))&" and "&trim(request("ExchangeDate1_h"))
		else
			strDCISQL=" where to_char(ExchangeDate,'hh') between "&trim(request("ExchangeDate_h"))&" and "&trim(request("ExchangeDate1_h"))
		end if
	end if
	'if request("Sys_BillUnitID")<>"" then
		'if strwhere<>"" then
			'strwhere=strwhere&" and a.BillUnitID='"&request("Sys_BillUnitID")&"'"
		'else
			'strwhere=" and a.BillUnitID='"&request("Sys_BillUnitID")&"'"
		'end if
	'end if
	if request("Sys_BillUnitID")<>"" and trim(request("Sys_BillNo"))="" then
		if request("Sys_BillMem")<>"" then
			if strDCISQL<>"" then
				strDCISQL=strDCISQL&" and RecordMemberID ="&request("Sys_BillMem")
			else
				strDCISQL=" where RecordMemberID="&request("Sys_BillMem")
			end if
		else
			if strDCISQL<>"" then
				strDCISQL=strDCISQL&" and RecordMemberID in(select MemberID from MemberData where UnitID in('"&request("Sys_BillUnitID")&"'))"
			else
				strDCISQL=" where RecordMemberID in(select MemberID from MemberData where UnitID in('"&request("Sys_BillUnitID")&"'))"
			end if
		end if
	end if
	if request("Sys_BillTypeID")<>"" then
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and BillTypeID='"&request("Sys_BillTypeID")&"'"
		else
			strDCISQL=" where BillTypeID='"&request("Sys_BillTypeID")&"'"
		end if
	end if

	if request("Sys_DCIReturnStatus")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and d.DCIReturnStatus "&request("Sys_DCIReturnStatus")
		else
			strwhere=" and d.DCIReturnStatus "&request("Sys_DCIReturnStatus")
		end if
	end if

	if trim(request("Sys_BillNo"))<>"" and trim(request("Sys_BillNo1"))<>"" then
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and BillSN in(select sn from billbase where BillNo between '"&Ucase(request("Sys_BillNo"))&"' and '"&Ucase(request("Sys_BillNo1"))&"') "
		else
			strDCISQL=" where BillSN in(select sn from billbase where BillNo between '"&Ucase(request("Sys_BillNo"))&"' and '"&Ucase(request("Sys_BillNo1"))&"') "
		end if
	elseif  trim(request("Sys_BillNo"))<>"" then
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and BillSN in(select sn from billbase where BillNo='"&Ucase(request("Sys_BillNo"))&"')"
		else
			strDCISQL=" where BillSN in(select sn from billbase where BillNo='"&Ucase(request("Sys_BillNo"))&"')"
		end if
	end if

	if request("Sys_CarNo")<>"" then
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and CarNo='"&Ucase(request("Sys_CarNo"))&"'"
		else
			strDCISQL=" where CarNo='"&Ucase(request("Sys_CarNo"))&"'"
		end if
	end if

	if request("Sys_ExchangeTypeID")<>"" then
		'Sys_ExchangeTypeID=split(trim(request("Sys_ExchangeTypeID")),"_")
		If trim(request("Sys_ExchangeTypeID"))="3" or trim(request("Sys_ExchangeTypeID"))="4" or trim(request("Sys_ExchangeTypeID"))="5" or trim(request("Sys_ExchangeTypeID"))="7" or trim(request("Sys_ExchangeTypeID"))="Y" Then

			if strDCISQL<>"" then
				strDCISQL=strDCISQL&" and ExchangeTypeID='N' and ReturnMarkType='"&trim(request("Sys_ExchangeTypeID"))&"'"
			else
				strDCISQL=" where ExchangeTypeID='N' and ReturnMarkType='"&trim(request("Sys_ExchangeTypeID"))&"'"
			end if
		else
			if strDCISQL<>"" then
				strDCISQL=strDCISQL&" and ExchangeTypeID='"&trim(request("Sys_ExchangeTypeID"))&"'"
			else
				strDCISQL=" where ExchangeTypeID='"&trim(request("Sys_ExchangeTypeID"))&"'"
			end if
		end if
	end if
	if request("Sys_RecordMemberID")<>"" and trim(request("Sys_BillNo"))="" then
		'Sys_ExchangeTypeID=split(trim(request("Sys_ExchangeTypeID")),"_")
		if strDCISQL<>"" then
			strDCISQL=strDCISQL&" and RecordMemberID="&trim(request("Sys_RecordMemberID"))
		else
			strDCISQL=" where RecordMemberID="&trim(request("Sys_RecordMemberID"))
		end if
	end if
	orderwhere=" order by a.ExchangeDate"
end if
DB_Display=request("DB_Display")
if DB_Display="show" then
	if trim(strwhere&strDCISQL)<>"" then
		strwhereToPrintCarData=strwhere

		strSQL="select a.SN,a.BillSN,a.RecordDate,a.ReturnMarkType,a.FileName,a.DCIReturnStatusID,a.ExchangeTypeID,a.DciErrorCarData,a.DCIErrorIDdata,b.ChName,(case when a.BillNo is null then (select billno from billbase where sn=a.billsn) else a.billno end) BillNo,a.CarNo,a.BillTypeID,a.EXCHANGEDATE,a.RecordMemberID,a.seqNo,a.BatchNumber,c.Content as BillTypeName,d.DCIReturn,d.StatusContent,d.DCIRETURNSTATUS,e.DCIActionName,f.DCIreturn as CarErrorSN,f.StatusContent as CarErrorContent,g.DCIreturn as DCIErrorSN,g.StatusContent as DCIErrorContent from (select * from DCILog"&strDCISQL&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&strwhere&orderwhere

		set rsfound=conn.execute(strSQL)

		if sys_City="基隆市" then
			strSQL="select sum(cnt) cnt from (select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','V') "&strwhere&" union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') or usetool=8) "&strwhere&" union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			strSQL="select sum(cnt) cnt from (select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','T','n') "&strwhere&" union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','T','n') and usetool=8 "&strwhere&" union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if

		set chksuess=conn.execute(strSQL)

		filsuess=CDbl(chksuess("cnt"))
		chksuess.close

		strSQL="select sum(cnt) cnt from (select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN "&strwhere&" and ExchangeTypeID='E' and DCIReturnStatusID='n' union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN "&strwhere&" and ExchangeTypeID='W' and DCIReturnStatusID in ('S','d','e') union all select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN "&strwhere&" and ExchangeTypeID='N' and (DCIReturnStatusID='n' or DCIReturnStatusID='h'))"
		set chksuess=conn.execute(strSQL)

		filClose=cdbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and d.DCIRETURNSTATUS='-1' "&strwhere
		set chksuess=conn.execute(strSQL)

		fildel=CDbl(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=CDbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=CDbl(Dbrs("cnt"))
		Dbrs.close

		if sys_City="基隆市" then
			strCnt="select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from (select * from DCILog"&strDCISQL&") a,DCIReturnStatus d,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=h.SN and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','T','n') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt=CDbl(Dbrs("cnt"))
		Dbrs.close

		CarSum=0:LeasesCnt=0:chkstrCarNo=""

		if request("DB_Selt")="BatchSelt" then
			strCnt="select count(*) as cnt from (select billno,carno,illegalspeed from BillBase where sn in(select billsn from DCILog"&strDCISQL&") and RecordStateID=0) a, (select BillNo,CarNo,DciReturnCarType from BilLBaseDciReturn where billno in(select billno from DCILog"&strDCISQL&") and ExChangeTypeID='W' and Status='Y') b, CarSpeed c where a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.DciReturnCarType=c.ID and a.IllegalSpeed>c.value"

			'strCnt="select count(*) as cnt from BillBase a, (select BillNo,CarNo,DciReturnCarType from BilLBaseDciReturn where billno in(select billno from DCILog"&strDCISQL&") and ExChangeTypeID='W' and Status='Y') b, CarSpeed c,UnitInfo d,MemberData e where a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.DciReturnCarType=c.ID and a.RecordMemberID=e.MemberID and a.BillUnitID=d.UnitID and a.RecordStateID=0 and a.IllegalSpeed>c.value and a.SN in(select distinct a.BillSN from (select * from DCILog"&strDCISQL&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&strwhere&")"

			set Dbrs=conn.execute(strCnt)
			CarSum=CDbl(Dbrs("cnt"))
			Dbrs.close



			if instr(request("Sys_BatchNumber"),"W")>0 or instr(request("Sys_BatchNumber"),"A")>0 then
				strCnt="select CarNo from billbasedcireturn where exchangetypeid='A' and ((mdrivername is not null and mdrivername<>owner) or (mdrivername is null and rdrivername is not null and rdrivername<>owner)) and exists(select 'Y' from dcilog A1 where billsn in(select billsn from dcilog where batchnumber in('"&Sys_BatchNumber&"')) and ExchangetypeID='A' and dcireturnstatusid='S' and carno=billbasedcireturn.carno) and carno in(select carno from dcilog where batchnumber in('"&Sys_BatchNumber&"'))"

				set Dbrs=conn.execute(strCnt)
				LeasesCnt=0
				while Not Dbrs.eof

					If trim(chkstrCarNo) <>"" Then chkstrCarNo=chkstrCarNo&"@"

					chkstrCarNo=chkstrCarNo&trim(Dbrs("CarNo"))

					LeasesCnt=LeasesCnt+1

					Dbrs.movenext
				wend
				Dbrs.close
			end If 
		end If 


		tmpSQL=strwhere&orderwhere

		chk_upprint=""
		if sys_City="高雄市" or sys_City="高港局" or sys_City="台中市" then
			if instr(request("Sys_BatchNumber"),"W")>0 then
				strSQL="select nvl(count(*),0) cnt from BillPrintJob where BatchNumber='"&trim(request("Sys_BatchNumber"))&"' and BillNo is null"
				set rschk = conn.execute(strSQL)
				if cdbl(rschk("cnt"))>0 then chk_upprint="true"
				rschk.close
			end if
		end If 
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>

<body>

<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF"><span class="style3">DCI 資料交換紀錄</span>(逕舉手開單入案後，請確認自動帶回之應到案處所是否與舉發單上相同)</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
					<%'If sys_City="彰化縣" or sys_City="台南市" then%>
						<span style="color: #FF0000;font-size: 18px;"><strong>※入案異常案件，請先參閱下方"異常案件處理方式"處理。</strong></span>
						<br />
					<%'end if%>
					<table cellspacing="0" border="1" style="border:1px #33CC33 solid;"  width="1200">
					<tr>
					<td height="50px">
						作業批號 
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							
							'這裡設定設定DCI Log 哪些縣市 批號要顯示幾天
							if sys_City="雲林縣" then
								nowdate=-2
							elseif sys_City="基隆市" then
								nowdate=-3
							else
								nowdate=-5
							end if
							strSQL="select Max(ExchangeDate) ExchangeDate,BatchNumber from DCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",nowdate, date)&" 00:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') group by BatchNumber order by ExchangeDate DESC"

										
							set rs=conn.execute(strSQL)
							cut=0
							while Not rs.eof
								ExchangeDate=gInitDT(trim(rs("ExchangeDate")))
		
								response.write "<option value="""&trim(rs("BatchNumber"))&""">"
								response.write ExchangeDate& " - "&cut&"　"&trim(rs("BatchNumber"))
								response.write "</option>"


								cut=cut+1
								rs.movenext
							wend
							rs.close
						%>
						</select>
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(trim(request("Sys_BatchNumber")))%>" size="20" maxlength="25">
						　結果
						<select name="Sys_DCIReturnStatus_Batch" class="btn1">
							<option value="">全部</option>
							<option value="is null"<%if trim(request("Sys_DCIReturnStatus_Batch"))="is null" then response.write " Selected"%>>未處理</option>
							<option value="=1"<%if trim(request("Sys_DCIReturnStatus_Batch"))="=1" then response.write " Selected"%>>正常</option>
							<option value="=-1"<%if trim(request("Sys_DCIReturnStatus_Batch"))="=-1" then response.write " Selected"%>>異常</option>
						</select>
						&nbsp;&nbsp;&nbsp;&nbsp;
						　車況
						<select name="Sys_CarType" class="btn1">
							<option value="">全部</option>
							<option value="1"<%if trim(request("Sys_CarType"))="1" then response.write " Selected"%>>主要駕駛</option>
						</select>　
						<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funSelt('BatchSelt');">&nbsp;<%
							'response.write "<input type=""button"" name=""Del"" value=""整批刪除"" onclick=""funDel('delall');"""
							'if Not CheckPermission(233,4) then response.write " disabled"
							'response.write ">"
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
							response.write "<input type=""button"" name=""ReSend"" value=""該批異常或未處理資料再次上傳"" style=""font-size: 18px;width:300px;height:25px;"" onclick=""funReSend('ReSend');"""
							if DBsum>0 and trim(request("Sys_BatchNumber"))<>"" then
								if Not ifnull(rsfound("RecordMemberID")) then
									if trim(rsfound("RecordMemberID"))<>trim(Session("User_ID")) and trim(Session("Credit_ID"))<>"A000000000" and trim(Session("Credit_ID"))<>"TIFFANY" and trim(Session("Credit_ID"))<>"19870107" then response.write " disabled"
								else
									response.write " disabled"
								end if
							else
								response.write " disabled"
							end if
							response.write ">"
						%>　
						
						<a href="javascript:void(0)" onclick="window.open('DCICreateFileLog.asp','WebPage4','left=130,top=30,location=0,width=800,height=500,resizable=yes,scrollbars=yes');"><font  class="font10">-> 查詢上傳下載歷程紀錄</font></a>
						<!--
							<img src="space.gif" width="15" height="8"></img><a href="uploadtime.htm" target="_blank"><font  class="font12"> ** 查詢系統上傳檔案時間點 ** </font></a>
						-->
					</td>
					</tr>
					</table>
					</td>
				</tr>
				<!-- <tr>
					<td>
						<hr>
					</td>
				</tr> -->
				<tr>
					<td nowrap>
					<!--建檔日期
						<input name="RecordDate" type="text" class="btn1" value="<%=request("RecordDate")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
						~
						<input name="RecordDate1" type="text" class="btn1" value="<%=request("RecordDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');"> 
						單位
						<select name="Sys_BillUnitID" class="btn1">
							<option value="">請選擇</option>
							<%strSQL="select UnitID,UnitName from UnitInfo order by UnitID"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("UnitID")&""""
								if rs1("UnitID")=request("Sys_BillUnitID") then response.write " selected"
								response.write ">"&rs1("UnitName")&"</option>"
								rs1.movenext
							wend 
							rs1.close%>
						</select>-->
					<table cellspacing="0" border="1" style="border:1px #33CC33 solid;"  width="1200">
					<tr>
					<td height="90px">
						上傳日期
						<input name="ExchangeDate" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("ExchangeDate"))
							else
								response.write gInitDT(DateAdd("d",-5, date))
							end if%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('ExchangeDate');">
						~
						<input name="ExchangeDate1" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("ExchangeDate1"))
							else
								response.write gInitDT(date)
							end if%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('ExchangeDate1');">
						時段
						<input name="ExchangeDate_h" type="text" class="btn1" value="<%=request("ExchangeDate_h")%>" size="1" maxlength="2" onkeyup="value=value.replace(/[^\d]/g,'')">
						時 ~ 
						<input name="ExchangeDate1_h" type="text" class="btn1" value="<%=request("ExchangeDate1_h")%>" size="1" maxlength="2" onkeyup="value=value.replace(/[^\d]/g,'')">時
						<img src="space.gif" width="5" height="10">
						<!--上傳人
						<select name="Sys_RecordMemberID" class="btn1">
							<option value="">全部</option>
							<%strSQL="select ChName,MemberID from MemberData where RecordStateID=0 and LeaveJobDate is null order by ChName"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("MemberID")&""""
								if DB_Display="show" then
									if trim(request("Sys_RecordMemberID"))=trim(rs1("MemberID")) then response.write " selected"
								else
									if trim(Session("User_ID"))=trim(rs1("MemberID")) then response.write " selected"
								end if
								response.write ">"&rs1("ChName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						<img src="space.gif" width="5" height="10">-->
						舉發單類別
						<select name="Sys_BillTypeID" class="btn1">
							<option value="">全部</option>
							<%strSQL="select * from DCIcode where TypeID=2 order by ID"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(request("Sys_BillTypeID"))=trim(rs1("ID")) then response.write " selected"
								response.write ">"&rs1("Content")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						DCI作業
						<select name="Sys_ExchangeTypeID" class="btn1">
							<option value="">全部</option>
							<%strSQL="select distinct DCIActionID,DCIActionName from DCIReturnStatus where DCIRETURNSTATUS=1"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								if trim(rs1("DCIActionID"))="N" then
									response.write "<option value=""3"""
									if trim(request("Sys_ExchangeTypeID"))="3" then response.write " selected"
									response.write ">單退</option>"

									response.write "<option value=""4"""
									if trim(request("Sys_ExchangeTypeID"))="4" then response.write " selected"
									response.write ">寄存</option>"

									response.write "<option value=""5"""
									if trim(request("Sys_ExchangeTypeID"))="5" then response.write " selected"
									response.write ">公示</option>"

									response.write "<option value=""Y"""
									if trim(request("Sys_ExchangeTypeID"))="Y" then response.write " selected"
									response.write ">撤銷</option>"

									response.write "<option value=""7"""
									if trim(request("Sys_ExchangeTypeID"))="7" then response.write " selected"
									response.write ">收受</option>"
								else
									response.write "<option value="""&rs1("DCIActionID")&""""
									if trim(request("Sys_ExchangeTypeID"))=trim(rs1("DCIActionID")) then response.write " selected"
									response.write ">"&rs1("DCIActionName")&"</option>"
								end if
								rs1.movenext
							wend
							rs1.close%>
							<!--<option value="1_A"<%if trim(request("Sys_ExchangeTypeID"))="1_A" then response.write " selected"%>>查詢</option>
							<option value="2_W"<%if trim(request("Sys_ExchangeTypeID"))="2_W" then response.write " selected"%>>入案</option>
							<option value="3_N"<%if trim(request("Sys_ExchangeTypeID"))="3_N" then response.write " selected"%>>退件</option>
							<option value="4_N"<%if trim(request("Sys_ExchangeTypeID"))="4_N" then response.write " selected"%>>寄存</option>
							<option value="5_N"<%if trim(request("Sys_ExchangeTypeID"))="5_N" then response.write " selected"%>>公示</option>
							<option value="6_E"<%if trim(request("Sys_ExchangeTypeID"))="6_E" then response.write " selected"%>>刪除</option>-->
						</select>
						<img src="space.gif" width="3" height="10">
						結果
						<select name="Sys_DCIReturnStatus" class="btn1">
							<option value="">全部</option>
							<option value="is null"<%if trim(request("Sys_DCIReturnStatus"))="is null" then response.write " Selected"%>>未處理</option>
							<option value="=1"<%if trim(request("Sys_DCIReturnStatus"))="=1" then response.write " Selected"%>>正常</option>
							<option value="=-1"<%if trim(request("Sys_DCIReturnStatus"))="=-1" then response.write " Selected"%>>異常</option>
						</select>
						<br>
						上傳單位
						<%=SelectUnitOption("Sys_BillUnitID","Sys_BillMem")%>
						<img src="space.gif" width="3" height="10">
						上傳人
						<%=SelectMemberOption("Sys_BillUnitID","Sys_BillMem")%>
						<img src="space.gif" width="5" height="10">
						舉發單號
						<input name="Sys_BillNo" type="text" class="btn1" value="<%=request("Sys_BillNo")%>" size="10" maxlength="9">
						<!--~
						<input name="Sys_BillNo1" type="text" class="btn1" value="<%=request("Sys_BillNo1")%>" size="10" maxlength="9">-->
						<img src="space.gif" width="3" height="10">
						車號
						<input name="Sys_CarNo" type="text" class="btn1" value="<%=request("Sys_CarNo")%>" size="8">
						<img src="space.gif" width="5" height="9">
						<input type="button" name="btnSelt" value="進階查詢" class="btn3" style="width:60px; height:20px;" onclick="funSelt('Selt');">
						<input type="button" name="cancel" value="清除" class="btn3" style="width:40px; height:20px;" onClick="location='DCIExchangeQry.asp'">
						<br />
						<span style="color: #FF0000;font-size: 16px;"><strong>※如使用進階查詢條件請點選『進階查詢』按鈕。</strong></span>
					</td>
					</tr>
					</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#1BF5FF">
		資料交換紀錄
		每頁<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功(<%=filClose%>筆結案<% 

				Response.Write "、"&LeasesCnt&"筆租賃"

			%>) , <%=errCatCnt%> 筆無效  ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-CDbl(filsuess)-CDbl(fildel)-CDbl(deldata)-CDbl(errCatCnt)%>筆未處理. )</strong>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10">批號</th>
					
					<th class="font10" width="3%" nowrap>上傳日期</th>
					<th class="font10" nowrap>作業</th>
					
					<th class="font10" nowrap>結果</th>
					<th class="font10">訊息</th>					
					<th class="font10" width="3%" nowrap>上傳人員</th>
					<th class="font10" nowrap>類別</th>
					<th class="font10">舉發單號</th>
					<th class="font10">車號</th>
					<!--<th class="font10">交換時間</th>-->
					<th class="font10">應到案處所</th>
					<th class="font10" width="3%" nowrap>廠牌.顏色<br>車藉狀況</th>
					<!--<th class="font10" nowrap>顏色</th>-->
					
					<th class="font10">操作</th>
					<th  width="5%"><font size="1">上下載檔案.序號</font></th>
				</tr>
				<%
				if DB_Display="show" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end If 
					if Not rsfound.eof then rsfound.move DBcnt
					ReturnMarkType=split("3,4,5,Y,7",",")
					ReturnMarkName=Split("單退,寄存,公示,撤消,收受",",")
					chkTypeID=0:chkBillNo=""
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						response.write "<tr bgcolor='#FFFFFF'"
						lightbarstyle 0 
						response.write ">"

						CNum=""

						if instr(rsfound("BatchNumber"),"N")>0 then

							strSQL="select cnt from (select RowNum cnt,billsn,UserMarkDate from (select billsn,UserMarkDate from billmailhistory where billsn in (select BillSN from DCILog where BatchNumber='"&trim(rsfound("BatchNumber"))&"') order by UserMarkDate) order by UserMarkDate) where BillSN="&rsfound("BillSN")

							set dci=conn.execute(strSQL)
							if not dci.eof then CNum=dci("cnt")
							dci.close

						else
							strSQL="select cnt from (select RowNum cnt,BillSN from (select BillSN from DCILog where BatchNumber='"&trim(rsfound("BatchNumber"))&"' order by BillSN) order by BillSN) where BillSN="&rsfound("BillSN")

							set dci=conn.execute(strSQL)
							if not dci.eof then CNum=dci("cnt")
							dci.close
						end if

						response.write "<td class=""font10"" >"&rsfound("BatchNumber")&"&nbsp("&CNum&")"&"</td>"      '" "&hour(rsfound("RecordDate"))&"時
						response.write "<td class=""font10"" nowrap>"&gInitDT(trim(rsfound("ExchangeDate")))&"</td>"
						'&"<br>"&right("00"&hour(trim(rsfound("ExchangeDate"))),2)&right("00"&minute(trim(rsfound("ExchangeDate"))),2)
						if trim(rsfound("ExchangeTypeID"))="N" then
							response.write "<td class=""font10"" align=""center"">"
							for arr=0 to Ubound(ReturnMarkType)
								if trim(ReturnMarkType(arr))=trim(rsfound("ReturnMarkType")) then
									response.write ReturnMarkName(arr)
									exit for
								end if
							next
							if arr>Ubound(ReturnMarkType) then response.write "送達註記"
							response.write "&nbsp;</td>"
						else
							response.write "<td class=""font10"" align=""center"" >"&rsfound("DCIActionName")&"</td>"
						end if
						
						if trim(rsfound("DCIRETURNSTATUS"))="1" then
							response.write "<td class=""font10"" nowrap>正常</td>"
						elseif trim(rsfound("DCIRETURNSTATUS"))="-1" then
							response.write "<td class=""font10"" nowrap><font color=""red"">異常</font></td>"
						else
							response.write "<td class=""font10"" nowrap>未處理</td>"
						end If 

						DCIerror="":dciSQL=""
						if trim(rsfound("DCIReturnStatusID"))="00" then
							if trim(rsfound("DciErrorCarData"))<>"" then
								dciSQL="'"&rsfound("DciErrorCarData")&"'"
							end if
							if trim(rsfound("DCIErrorIDdata"))<>"" then
								if trim(dciSQL)<>"" then
									dciSQL=dciSQL&",'"&rsfound("DCIErrorIDdata")&"'"
								else
									dciSQL="'"&rsfound("DCIErrorIDdata")&"'"
								end if
							end if
							if trim(dciSQL)<>"" then
								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DCIActionID='"&rsfound("ExchangeTypeID")&"E' and DCIReturn in("&dciSQL&")"
								set rsdci=conn.execute(strSQL)
								while Not rsdci.eof
									if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
									DCIerror=trim(DCIerror)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
									rsdci.movenext
								wend
								rsdci.close
							end if
						end if
						if trim(rsfound("BillTypeID"))="2" then
							strSQL="select ID,Content from DCICode where TypeID=10 and ID in(Select Rule4 from BillBaseDCIReturn where BillNo='"&rsfound("BillNo")&"' and CarNo='"&rsfound("CarNo")&"')"

							set rsdci=conn.execute(strSQL)
							while Not rsdci.eof
								if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
								DCIerror=trim(DCIerror)&rsdci("ID")&". "&rsdci("Content")
								rsdci.movenext
							wend
							rsdci.close
						end if

						Message=rsfound("DCIReturn")&". "&rsfound("StatusContent")
						'if trim(DCIerror)<>"" then Message=Message&"<br>"&DCIerror
						if trim(rsfound("CarErrorSN"))<>"" then Message=Message&"<br>"&rsfound("CarErrorSN")&". "&rsfound("CarErrorContent")
						if trim(rsfound("DCIErrorSN"))<>"" then Message=Message&"<br>"&rsfound("DCIErrorSN")&". "&rsfound("DCIErrorContent")

						response.write "<td class=""font10"" nowrap>"
						response.write Message
						response.write "</td>"
												
						'--------------------------------------------------------------
						response.write "<td class=""font10"" >"&rsfound("ChName")&"</td>"
						response.write "<td class=""font10"" >"&rsfound("BillTypeName")&"</td>"
						response.write "<td class=""font10"" >"&rsfound("BillNo")&"</td>"

						If i = (DBcnt+1) Then
							If not ifnull(rsfound("BillNo")) Then
								strSQL="select BillTypeID from billbase where billno='"&trim(rsfound("BillNo"))&"'"
								set chktype=conn.execute(strSQL)
								If not chktype.eof Then
									chkTypeID=cdbl(chktype("BillTypeID"))
									chkBillNo=trim(rsfound("BillNo"))
								End if
								chktype.close
							End if
						End If 
						
						chkRdCar=rsfound("CarNo"):chkRdCarColor=""

						If instr("@"&chkstrCarNo&"@","@"&trim(rsfound("CarNo"))&"@")>0 then

							chkRdCarColor=" style=""color:#ff0000;"""
							chkRdCar="<b>(長租)"&rsfound("CarNo")&"</b>"
						End if 
						
						response.write "<td class=""font10"" "&chkRdCarColor&" nowrap>"&chkRdCar&"</td>"
						'response.write "<td class=""font10"">"&rsfound("EXCHANGEDATE")&"</td>" '交換時間
						
						
						'--------------------------------------------------------------
						'StrBass="select  b.Content as CarTypeName,c.Content as CarColor,d.Content as Rule4Name from BillBaseDCIReturn a,(select ID,Content from DCICode where TypeID=5) b,(select ID,Content from DCICode where TypeID=4) c,(select ID,Content from DCICode where TypeID=10) d where a.DciReturnCarType=b.ID(+) and a.DciReturnCarColor=c.ID(+) and a.Rule4=d.ID(+) and a.BillNo='"&rsfound("BillNo")&"' and a.CarNo='"&rsfound("CarNo")&"'"

						StrBass="select a.A_Name,a.DciReturnCarColor,c.ID as CarStatusID,c.Content as CarStatusName,d.ID as Rule4,d.Content as Rule4Name,e.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='A'  and CarNo='"&rsfound("CarNo")&"') a,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+)"
						set rsCarType=conn.execute(strBass)
						Sys_DciReturnCarColor="":Sys_DCIStationName="":Sys_A_Name="":Sys_CarStatusID="":Sys_CarStatusName="":Sys_Rule4="":Sys_Rule4Name="":Sys_CarColorID="":Sys_CarColorName=""
						if not rsCarType.eof then
							Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
							Sys_DCIStationName=trim(rsCarType("DCIStationName"))
							Sys_A_Name=trim(rsCarType("A_Name"))
							Sys_CarStatusID=trim(rsCarType("CarStatusID"))
							Sys_CarStatusName=trim(rsCarType("CarStatusName"))
							Sys_Rule4=trim(rsCarType("Rule4"))
							Sys_Rule4Name=trim(rsCarType("Rule4Name"))
						end if
						rsCarType.close

						StrBass="select a.DciReturnCarColor,b.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='W' and CarNo='"&trim(rsfound("CarNo"))&"' and BillNo='"&trim(rsfound("BillNo"))&"') a,Station b where a.DCIReturnStation=b.DCIStationID(+)"

						set rsCarType=conn.execute(strBass)
						if not rsCarType.eof then
							Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
							If trim(rsfound("ExchangeTypeID"))<>"A" then Sys_DCIStationName=trim(rsCarType("DCIStationName"))
						end if
						rsCarType.close

						if len(Sys_DciReturnCarColor)>1 then Sys_DciReturnCarColor=left(Sys_DciReturnCarColor,1)&","&right(Sys_DciReturnCarColor,1)
						if ifnull(Sys_DciReturnCarColor) then Sys_DciReturnCarColor=""
						Sys_CarColorID=split(Sys_DciReturnCarColor,",")
						for y=0 to ubound(Sys_CarColorID)
							strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
							set rscolor=conn.execute(strColor)
							if not rscolor.eof then
								if trim(Sys_CarColorName)<>"" then Sys_CarColorName=Sys_CarColorName&","
								Sys_CarColorName=Sys_CarColorName&trim(rscolor("Content"))
							end if
							rscolor.close
						next
						response.write "<td class=""font10""  nowrap>"&Sys_DCIStationName&"</td>"
						response.write "<td class=""font10"" nowrap > "&Sys_A_Name
						if trim(Sys_A_Name)<>"" then response.write ". "
						response.write Sys_CarColorName
						response.write "<br>"

						'response.write "<td class=""font10"" nowrap>"&rsCarType("CarColor")&"</td>"
						if not ifnull(Sys_CarStatusID) then response.write Sys_CarStatusID&"_"&Sys_CarStatusName

						if not ifnull(Sys_CarStatusID) and Not ifnull(Sys_Rule4) then response.write "<br>"
						if not ifnull(Sys_Rule4) then response.write Sys_Rule4&"_"&Sys_Rule4Name
						response.write "</td>"
					
						'--------------------------------------------------------------
						response.write "<td class=""font10"">"
						'response.write "<input type=""button"" name=""Update"" value=""詳細資料"" onclick=""funDataDetail('"&rsfound("BillSN")&"');"""
						'response.write ">"

						if trim(Session("Credit_ID"))="A000000000" then 
							if trim(rsfound("BillTypeName"))="逕舉" Then

								response.write "<input type=""button"" name=""Update"" value=""修改"" class=""btn3"" style=""width:40px; height:20px;"" onclick=""funUpdate('"&rsfound("BillSN")&"');"""
								response.write ">"

							else

								response.write "<input type=""button"" name=""Update"" value=""修改"" class=""btn3"" style=""width:40px; height:20px;"" onclick=""funUpdate2('"&rsfound("BillSN")&"');"""
								response.write ">"

							end If
						end If 
												' 
						if (trim(rsfound("RecordMemberID"))=trim(Session("User_ID")) and trim(rsfound("DCIRETURNSTATUS"))="-1") or trim(Session("Credit_ID"))="A01" Then
							'攔停
							'高雄市攔停已結案
							StrBillStatus="select BillStatus,BillUnitID from billbase where sn="&trim(rsfound("BillSN"))
							Set rsBillStatus=conn.execute(strBillStatus)
							If not rsBillStatus.eof Then
								If Not isnull(rsfound("DCIReturnStatusID")) then
									if (sys_City="高雄市" or sys_City="高港局") and trim(rsBillStatus("BillStatus"))="9" and (trim(rsBillStatus("BillUnitID"))="0861" or trim(rsBillStatus("BillUnitID"))="0862" or trim(rsBillStatus("BillUnitID"))="0863" or trim(rsBillStatus("BillUnitID"))="0864" or trim(rsBillStatus("BillUnitID"))="0871") Then %>

										<input type="button" name="b1" value="修改" class="btn3" style="width:40px; height:20px;" onclick='window.open("../BillKeyIn/BillKeyIn_TakeCar_Update.asp?BillSN=<%=trim(rsfound("BillSN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
										'1:查詢 ,2:新增 ,3:修改 ,4:刪除
										if CheckPermission(234,3)=false then
											response.write "disabled"
										end if

										%> style="font-size: 12pt; width: 45px; height:26px;"><%						
									elseif trim(rsfound("BillTypeName"))="逕舉" Then
										If sys_City<>"高雄市" then
											response.write "<input type=""button"" name=""Update"" value=""修改"" class=""btn3"" style=""width:40px; height:20px;"" onclick=""funUpdate('"&rsfound("BillSN")&"');"""
											if Not CheckPermission(233,3) then 
												response.write " disabled"
											elseif sys_City="保二總隊四大隊二中隊" and (session("unit_id")="A01" or session("unit_id")="B01" or session("unit_id")="C01") then
												response.write " disabled"
											end If 
											response.write ">"
										End If 
									else
										response.write "<input type=""button"" name=""Update"" value=""修改"" class=""btn3"" style=""width:40px; height:20px;"" onclick=""funUpdate2('"&rsfound("BillSN")&"');"""
										if Not CheckPermission(233,3) then 
											response.write " disabled"
										elseif sys_City="保二總隊四大隊二中隊" and (session("unit_id")="A01" or session("unit_id")="B01" or session("unit_id")="C01") then
											response.write " disabled"
										end If 
										response.write ">"

									end If
								End If 
							end if
							rsBillStatus.close
						end if
						'高雄市 派出所不能點強制入案(trim(rsfound("RecordMemberID"))=trim(Session("User_ID")) and sys_City<>"高雄市") or trim(Session("Credit_ID"))="A000000000" Or (sys_City="台中市" And trim(Session("Credit_ID"))="Z016") or (sys_City="高雄市" and trim(Session("UnitLevelID"))="1" and trim(rsfound("RecordMemberID"))=trim(Session("User_ID"))) or (sys_City="高雄市" and trim(Session("UnitLevelID"))="2")
						if (trim(rsfound("RecordMemberID"))=trim(Session("User_ID"))) or trim(Session("Credit_ID"))="A000000000" Or (sys_City="台中市" And trim(Session("Credit_ID"))="Z016") then
							if (trim(rsfound("DciErrorCarData"))="V" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DCIReturnStatusID"))="L" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DCIReturnStatusID"))="x" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="4" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DCIReturnStatusID"))="n" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="F" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="n" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="o" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="r" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="s" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="G" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="W" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="#" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="t" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="w" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="+" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="-" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="X" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciErrorCarData"))="C" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciReturnStatusID"))="00" and trim(rsfound("ExchangeTypeID"))="W") or (trim(rsfound("DciReturnStatusID"))="N" and trim(rsfound("ExchangeTypeID"))="W") Or (trim(Session("Credit_ID"))="A000000000" and trim(rsfound("ExchangeTypeID"))="W")then
								response.write "<input type=""button"" name=""Update"" value=""強制入案(僅供列印各式清冊，監理站不會有資料)"" class=""btn3"" style=""width:300px; height:20px;"" onclick=""funUpdate3('"&rsfound("BillSN")&"');"">"
							ElseIf (trim(rsfound("ExchangeTypeID"))="E" And Session("Credit_ID")="A000000000") Then
								response.write "<input type=""button"" name=""Update"" value=""強制刪除"" class=""btn3"" style=""width:80px; height:20px;"" onclick=""funDelete3('"&rsfound("BillSN")&"');"">"
							elseif (trim(rsfound("ExchangeTypeID"))="E" and trim(rsfound("DCIReturnStatusID"))="N") or (trim(rsfound("ExchangeTypeID"))="E" and trim(rsfound("DCIReturnStatusID"))="B") Then
								'有做過強制入案的才可用強制刪除
								strDel="select * from DCISTATUSUPDATE where Billsn="&rsfound("BillSN")
								set rsDel=conn.execute(strDel)
								if not rsDel.eof then
									response.write "<input type=""button"" name=""Update"" value=""強制刪除"" class=""btn3"" style=""width:80px; height:20px;"" onclick=""funDelete3('"&rsfound("BillSN")&"');"">"
								end if
								rsDel.close
								set rsDel=Nothing
							
							elseif (trim(rsfound("ExchangeTypeID"))="N" and trim(rsfound("DCIReturnStatusID"))="n") Or (trim(rsfound("ExchangeTypeID"))="N" and sys_City="台中市" And trim(Session("Credit_ID"))="Z016") Or (trim(rsfound("ExchangeTypeID"))="N" and sys_City="南投縣") Or (trim(rsfound("ExchangeTypeID"))="N" and trim(Session("Credit_ID"))="A000000000") Then

								response.write "<input type=""button"" name=""Update"" value=""結案轉成功"" class=""btn3"" style=""width:90px; height:20px;"" onclick=""funUpdate4('"&rsfound("SN")&"','"&rsfound("ReturnMarkType")&"');"">"
			
							end if
 						end If
						If sys_City="嘉義市" or sys_City="苗栗縣" or sys_City="嘉義縣" or (sys_City="高雄市" And Session("Credit_ID")="A000000000") Then
						'要新增縣市的話要新增TABLE
						'CREATE TABLE TRAFFIC.ReloadReason(
						'	DCISN number,
						'	Reason varchar2(200),
						'	RecordMemberID number,
						'	RecordDate date
						')
							If (trim(rsfound("ExchangeTypeID"))="N" And (trim(rsfound("RETURNMARKTYPE"))="5" Or trim(rsfound("RETURNMARKTYPE"))="4" Or trim(rsfound("RETURNMARKTYPE"))="7") And (trim(rsfound("DCIReturnStatusID"))="S" or trim(rsfound("DCIReturnStatusID"))="n")) and (trim(Session("Credit_ID"))="A01" or trim(Session("Credit_ID"))="A000000000" or trim(rsfound("RecordMemberID"))=trim(Session("User_ID")))  Then
								response.write "<input type=""button"" name=""Update"" value=""重新上傳"" class=""btn3"" style=""width:80px; height:20px;"" onclick=""funReUpload('"&rsfound("SN")&"');"">"

								'response.write "<input type=""button"" name=""Update3"" value=""另案舉發"" class=""btn3"" style=""width:60px; height:20px;"" onclick=""funReOther('"&rsfound("SN")&"');"">"
								
							End If 
							If sys_City="嘉義市" Or sys_City="高雄市" then
								strReUp="select a.RecordDate,a.Reason,b.ChName from ReloadReason a,Memberdata b where a.DciSN="&Trim(rsfound("SN"))&" and a.RecordMemberID=b.MemberID"
								Set rsReUp=conn.execute(strReUp)
								If Not rsReUp.Bof Then rsReUp.MoveFirst 
								While Not rsReUp.Eof
									response.write "<br>"&Year(rsReUp("RecordDate"))-1911&Right("00"&Month(rsReUp("RecordDate")),2)&Right("00"&Day(rsReUp("RecordDate")),2)&"("&Trim(rsReUp("ChName")) &")"&":"&Trim(rsReUp("Reason"))
									rsReUp.MoveNext
								Wend
								rsReUp.close
								Set rsReUp=Nothing 
							end if
						End If 
'						if trim(rsfound("RecordMemberID"))=trim(Session("User_ID")) and isnull(trim(rsfound("DCIRETURNSTATUS")))  and isnull(rsfound("FileName"))  then
'							response.write "<input type=""button"" name=""Del"" value=""不上傳"" onclick=""funDel('"&rsfound("SN")&"');"""
'							if Not CheckPermission(233,4) then response.write " disabled"
'							response.write ">"
'						end if
						response.write "&nbsp;</td>"
						'--------------------------------------------------------------
						response.write "<td font size=""1"" nowrap>"
						if trim(rsfound("DCIReturnStatusID"))<>"" then
							response.write "<a href='DCIfile.asp?DCIfile=/UP/"&trim(rsfound("FileName"))&"' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&"</font>&nbsp;<font size='1' color=""Red"">"&trim(rsfound("seqNo"))&"</font></a><br>"
							response.write "<a href='DCIfile.asp?DCIfile=/Down/"&trim(rsfound("FileName"))&".big' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&".big </font>&nbsp;<font size=""1"" color=""Red"">"&trim(rsfound("seqNo"))&"</font></a>"
						else
							response.write "<a href='DCIfile.asp?DCIfile=/UP/"&trim(rsfound("FileName"))&"' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&"</font>&nbsp;<font size=""1""  color=""Red"">"&trim(rsfound("seqNo"))&"</font></a><br>"
							response.write "<font size='1'>" & trim(rsfound("FileName"))& "&nbsp;"&trim(rsfound("seqNo"))&"</font>"
						end if
						response.write "</td>"
						response.write "</tr>"
						response.flush
						rsfound.movenext
					next
				end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="30" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<input type="button" name="MoveFirst" value="第一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(0);">
			<input type="button" name="MoveUp" value="上一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(10);">
			<input type="button" name="MoveDown" value="最後一頁" class="btn3" style="width:60px; height:20px;" onclick="funDbMove(999);">
			<br>

		<%if sys_City="高雄市" then%>
			<input type="button" name="Submit42342funBillCheck22" value="舉發單核對清冊" class="btn3" style="width:110px; height:25px;" onclick="funBillCheck()">
		<%end if%>
		<%if sys_City="台東縣" Or sys_City="宜蘭縣" Or sys_City="南投縣" Or sys_City="花蓮縣" Or sys_City="苗栗縣" Or sys_City="嘉義縣" Or sys_City="屏東縣" Or sys_City="基隆縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="金門縣" Or sys_City="高雄市" Or sys_City="台南市" Or sys_City="雲林縣" Or sys_City="保二總隊四大隊二中隊" or trim(Session("Credit_ID"))="A000000000" then%>
			<input type="button" name="Submit4234222" value="<%
			If sys_City="彰化縣" Then
				response.write "民眾檢舉車籍資料檢核"
			Else
				response.write "數位影像車籍資料"
			End if			
			%>" class="btn3" style="width:130px; height:25px;" onclick="funImgCarDataList()">
		<%end if%>

		<%if sys_City="花蓮縣" then%>
			<input type="button" name="Submit4234222" value="入案清冊(地址審核)" class="btn3" style="width:110px; height:25px;" onclick="funDCILogCarListDetail()">
			<input type="button" name="Submit4234222" value="車籍清冊(地址審核)" class="btn3" style="width:110px; height:25px;" onclick="funStopCarListDetail()">
		<%end if%>
		<%if sys_City="台東縣" then%>
			<input type="button" name="Submit4234222" value="列印催繳比對資料" class="btn3" style="width:110px; height:25px;" onclick="funchgCkeckDataList()">
		<%End If %>

		<%if sys_City="苗栗縣" then%>
			<input type="button" name="Submit4234222" value="車籍清冊(地址審核)" class="btn3" style="width:110px; height:25px;" onclick="funStopCarListDetail()">
			<input type="button" class="btn3" name="Submit4232" value="入案資料審核清冊(苗)" onclick="funBillCaseDataDetail();">

			<input type="button" name="btnExecel" value="公告匯出" class="btn3" style="width:70px; height:25px;" onclick="funOpenGovList();">
		<%end if%>

		<%if sys_City="苗栗縣" or sys_City="台中市" then%>
			
			<input type="button" class="btn3" name="btnExecel" value="戶籍地址補正清冊" onclick="funStoreAndSendAddress();">

		<%end if%>

			<input type="button" name="Submit4234222" value="列印車籍資料" class="btn3" style="width:120px; height:25px;" onclick="funchgCarDataList()">

			<input type="button" name="CarSpeed" value="稽核特殊車種車速(<%=CDbl(CarSum)%>)" class="btn3" style="width:120px; height:25px;" onclick="funCarDetail()">

			<input type="button" name="btnExecel" value="轉換成Excel" class="btn3" style="width:120px; height:25px;" onclick="funchgExecel();">
		<br><%
			if sys_City="高雄市" or sys_City="高港局" or sys_City="台中市" or sys_City="嘉義縣" or (sys_City="彰化縣" and trim(Session("Credit_ID"))="A000000000") then
				Response.Write "<br><span class=""style1"">至少要用批號查詢確認過一次才可以使用代印上傳。<br>連續單號請在中間加 '~'，跳印請用 ',' 區隔　"
				Response.Write "例如連續單號：GE0000001~GE0000006　"
				Response.Write "跳號：GE0000001,GE0000003,GE0000005</span>"

				response.write "<br><input type=""button"" name=""RePrint"" ID=""RePrint"" value=""該批資料代印上傳"" class=""btn3"" style=""width:110px; height:25px;"" onclick=""funRePrint();"""


				if Instr(request("Sys_BatchNumber"),"W")<=0 then
					Response.Write " disabled"

				elseIf  chkTypeID=1 or ifnull(chkBillNo) Then
					Response.Write " disabled"

				elseif filsuess=0 then
					Response.Write " disabled"

				End if
				
				Response.Write ">"

				Response.Write "　　　補印單號　<input name=""Sys_ReBillNo"" type=""text"" class=""btn1"" value="""" size=""50""><br>"
				Response.Write "　　　　　　　備註　<input name=""Sys_Note"" maxlength=""20"" type=""text"" class=""btn1"" value="""" size=""40"">"
			end if
		%>
        <!--<span class="style3">
        DCI檔案名稱
        <input name="textfield42324" type="text" value="" size="14" maxlength="13">
        </span>-->
     
       
   <!--
		<%'if trim(request("Sys_ExchangeTypeID"))="W" then '入案%>
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="btnprintBill" value="列印違規通知單" onclick="funPrintStyle()">
		<input type="button" name="btnprintBill" value="列印送達證書" onclick="funUrgeStyle()">
		<%'end if%>

        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit42342" value="大宗掛號清冊" onclick="funMailList()">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">

		<%'if trim(request("Sys_ExchangeTypeID"))="W" then '入案%>
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit4234" value="逕舉移送清冊" onclick="funReportSendList()">
		<span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit4335" value="攔停移送清冊" onclick="funStopSendList()">
		<%'end if%>

        <span class="style3"><img src="space.gif" width="13" height="8"></span>
		<span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList()">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit43635" value="無效清冊" onclick="funUselessSendList()">
		<span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit43635" value="結案清冊" onclick="funCaseCloseSendList()">
		<%'if trim(request("Sys_ExchangeTypeID"))="N" then '退件%>
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
        <input type="button" name="Submit4233" value="退件清冊" onclick="funReturnSendList()">
		<%'end if
		'if trim(request("Sys_ExchangeTypeID"))="N" then '寄存%>
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
		<input type="button" name="Submit4233" value="寄存送達清冊" onclick="funStoreSendList()">
		<%'end if
		'if trim(request("Sys_ExchangeTypeID"))="N" then '公示%>
        <span class="style3"><img src="space.gif" width="13" height="8"></span>
		<input type="button" name="Submit4232" value="公示送達清冊" onclick="funGovSendList()">
		<%'end if%>
		-->
	</td>
	
  </tr>
</table>
<br><b>
* 系統上傳DCI 時間 :</b> 0850 ,  1050 ,  1250 , 1450 ,  1620 , 1850 。 請於各梯次 <b>前5分鐘</b> 上傳監理所。
<br>
<b>* DCI 抓取檔案時間  :</b> 0900 ,  1100 ,  1300 , 1500 ,  1630 , 1900 。 </b> 
<br>
<!--
* 上傳監理站不須等到前一批回傳，同一時段內可上傳多批檔案。 <br>
* 舉發單入案如經車籍查詢，則需要等車籍查詢結果回傳才可再次入案，<br>
��	系統也提供跳過車籍查詢，直接入案之功能。但此功能則無法提供廠牌功能。
-->
<br>
<b>* 入案回傳<font color="red">正常</font>案件欲修正 :�僉I選 " 刪除 " 按鈕。確認刪除成功後再次建檔入案。</b>
<br>
<b>* 入案回傳<font color="red">異常</font>案件欲修正 :�僉I選 " 修改 " 按鈕完成資料修正後，再點選  "該批異常資料再次上傳" 按鈕 ，重新上傳 <br><br> </b>

<b><font style="color: #FF0000;font-size: 20px;">異常案件處理方式：</font></b><br>
<b><font style="font-size: 20px;">※入案</font><br></b>
��	<b><font style="font-size: 20px;">未新增</font><br></b>
==>如該批資料多筆皆回傳 未新增 之異常訊息，通常為監理站線路或資料庫異常，<br>==>請執行該批異常資料再次上傳<br>

��	<b><font style="font-size: 20px;">時間不明確</font><br></b>
==>請檢違規日,應到案日是否輸入有問題，修正後執行該批異常資料再次上傳<br>
��	<b><font style="font-size: 20px;">條款錯誤</font><br></b>
==>請檢查條款是否正確、是否引用到舊法條，修正後再次上傳<br>
��	<b><font style="font-size: 20px;">證號重號剔退</font></b><br>
==>請補入生日後再次上傳<br>
 ��	<b><font style="font-size: 20px;">失竊註銷 (但查詢署端系統後發現車輛實際已經尋回)</font><br></b>
==>此類資料導因警政署與監理站資料更新不一致所引起，須手動移送監理站，<br>
==>系統會紀錄無法入案原因，隨時備查。<br>
��	<b><font style="font-size: 20px;">未知,找不到</font><br></b>
==>請確認車號、車種是否正確，如為攔停案件還需確認身份證號是否正確，修正後執行該批異常資料再次上傳<br>
��	<b><font style="font-size: 20px;">攔停車駕條款一起</font><br></b>
==>監理站規定舉發單不可將處車主法條及處駕駛法條開在同一張<br>
��	<b><font style="font-size: 20px;">處所不明</font><br></b>
==>如為酒駕案件，監理站規定應到案處所需輸入行為地所屬監理站<br>
��	<b><font style="font-size: 20px;">繳註銷後案、車主已死亡、報廢車輛等車駕籍異常</font><br></b>
==>車駕籍異常案件，監理站系統會剔退，無法由執法系統入案，需人工移交由監理站代為入案<br>
<br>
<br>
<b><font style="font-size: 20px;">※單退、送達註記</font><br></b>
 ��	<b><font style="font-size: 20px;">找不到此筆資料　或　無此車號/無此證號</font><br></b>
==>請向監理站確認單號和車號是否都正確<br>
==>如為強制入案案件，須確認是否有移送給監理站做建檔入案<br>
��	<b><font style="font-size: 20px;">已開裁決書</font><br></b>
==>案件狀態被監理站鎖定，請洽詢該案件所屬監理機關<br>
<br>
<br>
<b><font style="font-size: 20px;">※刪除</font><br></b>
 ��	<b><font style="font-size: 20px;">找不到此筆資料</font><br></b>
==>請向監理站確認單號和車號是否都正確<br>
==>如為強制入案案件，須確認是否有移送給監理站做建檔入案<br>
 ��	<b><font style="font-size: 20px;">已結案無法做刪除</font><br></b>
==>監理站已將此筆舉發單結案，無法透過線上系統刪除此筆案件，請洽詢監理站<br>



<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="PBillSN" value="<%
	if trim(request("PBillSN"))<>"" then
		response.write request("PBillSN")
	else
		response.write BillSN
	end if%>">
	<input type="Hidden" name="printStyle" value="">
	<input type="Hidden" name="Sys_MailDate" value="">
	<input type="Hidden" name="Sys_JudeAgentSex" value="">
	<input type="Hidden" name="chk_UpPrint" value="<%=chk_upprint%>">
	
	<input type="Hidden" name="Hidden_ReturnMarkType" value="">
</form>
<form Name=CarForm method="post">
<input type="Hidden" name="TempSQL" value="<%=strwhere%>">
<input type="Hidden" name="strDCISQL" value="<%=strDCISQL%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var winopen;
var sys_City='<%=sys_City%>';
<%
if trim(session("ManagerPower"))="1" and sys_City="苗栗縣" then
	response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"
elseif sys_City<>"苗栗縣" then
	response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"
end if
%>
function funSelt(DBKind){
	var error=0;
	if(DBKind=='Selt'){
		/*if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日輸入不正確!!");
			}
		}
		if (error==0){
			if(myForm.RecordDate1.value!=""){
				if(!dateCheck(myForm.RecordDate1.value)){
					error=1;
					alert("建檔日輸入不正確!!");
				}
			}
		}*/
		if (error==0){
			if(myForm.ExchangeDate.value!=""){
				if(!dateCheck(myForm.ExchangeDate.value)){
					error=1;
					alert("上傳日輸入不正確!!");
				}
			}
		}
		if (error==0){
			if(myForm.ExchangeDate1.value!=""){
				if(!dateCheck(myForm.ExchangeDate1.value)){
					error=1;
					alert("上傳日輸入不正確!!");
				}
			}
			if (error==0){
				myForm.PBillSN.value="";
				//CarForm.PBillSN.value="";
				myForm.DB_Move.value="";
				myForm.DB_Selt.value=DBKind;
				myForm.DB_Display.value='show';
				myForm.submit();
			}
		}
	}else if(DBKind=='BatchSelt'){
		myForm.PBillSN.value="";
		//CarForm.PBillSN.value="";
		myForm.DB_Move.value="";
		myForm.DB_Selt.value=DBKind;
		myForm.DB_Display.value='show';
		myForm.submit();
	}
}
function fnBatchNumber(){
	myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}
function funDataDetail(SN){
	UrlStr="ViewBillBaseData_Car.asp?BillSn="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funUpdate(SN){
	UrlStr="../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funUpdate2(SN){
	UrlStr="../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funReUpload(SN){
	UrlStr="ReUpdateSend.asp?SN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}

function funReOther(SN){
	UrlStr="OtherBillQry.asp";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
//強迫入案，把失竊註銷但要入案的案件改為入案正常
function funUpdate3(SN){
	UrlStr="CaseInStatus_Update.asp?BillSN="+SN;
	window.open(UrlStr,"WebPage_Detailfd","left=0,top=0,location=0,width=600,height=350,resizable=yes,scrollbars=yes,menubar=yes")
	//myForm.submit();
}
function funDelete3(SN){
	UrlStr="CaseInStatus_Delete.asp?BillSN="+SN;
	window.open(UrlStr,"WebPage_DetailDelete","left=0,top=0,location=0,width=600,height=350,resizable=yes,scrollbars=yes,menubar=yes")
	//myForm.submit();
}
function funUpdate4(SN,ReturnMarkType){
	myForm.SN.value=SN;
	myForm.Hidden_ReturnMarkType.value=ReturnMarkType;
	myForm.DB_state.value="CloseToSuss";
	myForm.submit();
}
function funDel(SN){
	if(SN=='delall'&&myForm.DB_Display.value!=""){
		if(myForm.Sys_BatchNumber.value!=""){
			myForm.SN.value=SN;
			myForm.DB_state.value="Del";
			myForm.submit();
		}else{
			alert("請先進行批號查詢!!");
		}
	}else if(myForm.DB_Display.value!=""){
		myForm.SN.value=SN;
		myForm.DB_state.value="Del";
		myForm.submit();
	}else{
		alert("請先進行批號查詢!!");
	}
}

function funStoreAndSendAddress(){
	if (myForm.Sys_BatchNumber.value==""){
		alert("請先輸入作業批號！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="BillBaseStoreAndSendAddressList.asp";
		myForm.action=UrlStr;
		myForm.target="BillBaseStoreAndSendAddressList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funReSend(SN){
	if(SN=='ReSend'&&myForm.DB_Display.value!=""){
		if(myForm.Sys_BatchNumber.value!=""){
			if(confirm("是否確定要再次上傳?")){
				myForm.SN.value="";
				myForm.DB_state.value="ReSend";
				myForm.submit();
			}
		}else{
			alert("請先進行批號查詢!!");
		}
	}else{
		alert("請先進行批號查詢!!");
	}
}
function funRePrint(){
	if(myForm.Sys_BatchNumber.value=='<%=request("Sys_BatchNumber")%>'){
		if(myForm.DB_Display.value!=""){
			if(myForm.Sys_BatchNumber.value!=""){
				myForm.SN.value="";
				var dt = new Date();
				if(myForm.chk_UpPrint.value!=""){
					if(confirm("該批資料已有上傳是否再次上傳?")){
						runServerScript("UpBatchNumber.asp?batchnumber="+myForm.Sys_BatchNumber.value+"&ReBillNo="+myForm.Sys_ReBillNo.value+"&Sys_Note="+myForm.Sys_Note.value+"&PrintCnt=<%=filsuess%>&nowtime="+dt);
					}
				}else{
					document.all.RePrint.disabled=true;
					runServerScript("UpBatchNumber.asp?batchnumber="+myForm.Sys_BatchNumber.value+"&ReBillNo="+myForm.Sys_ReBillNo.value+"&Sys_Note="+myForm.Sys_Note.value+"&PrintCnt=<%=filsuess%>&nowtime="+dt);

					if (sys_City=='高雄市'){

						UrlStr="http://10.133.2.176/traffic/Query/ImageReSize.asp";

						myForm.action=UrlStr;
						myForm.target="ImageReSize";
						myForm.submit();
						myForm.action="";
						myForm.target="";
					}
				}
			}else{
				alert("請先進行批號查詢!!");
			}
		}else{
			alert("請先進行批號查詢!!");
		}
	}else{
		alert("請先進行批號查詢!!");
	}
}
function funsubmit(){
	winopen.close();
	if(myForm.printStyle.value=='0'){
		/*window.parent.frames("mainFrame").location="BillPrints.asp";
		myForm.action="BillPrints.asp";*/
		UrlStr="BillPrints.asp";
	}else{
		/*window.parent.frames("mainFrame").location="BillPrints_a4.asp";
		myForm.action="BillPrints_a4.asp";*/
		UrlStr="BillPrints_a4.asp";
	}
	/*myForm.target="mainFrame";
	myForm.submit();
	myForm.action="";
	myForm.target="";*/
	newWin(UrlStr,"JudeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="JudeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}
function funUrgeList(){
	winopen.close();
	UrlStr="BillPrints_legal.asp";
	newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="UrgeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}
function funUrgeStyle(){

		UrlStr="UrgeStyle.asp";
		newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
		myForm.action="UrgeStyle.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funchgprint(){
	winopen.DP();
}
function funchgExecel(){
	CarForm.action="DCIExchangeQry_Execel.asp";
	CarForm.target="inputWin";
	CarForm.submit();
	CarForm.action="";
	CarForm.target="";
}
function funOpenGovList(){
	CarForm.action="OpenGovList_Execel.asp";
	CarForm.target="OpenGovList";
	CarForm.submit();
	CarForm.action="";
	CarForm.target="";
}
function funBillCheck(){
	CarForm.action="DCIBillCheck.asp";
	CarForm.target="inputWin";
	CarForm.submit();
	CarForm.action="";
	CarForm.target="";
}
function funCarDetail(){
		CarForm.action="CarSpeed.asp";
		CarForm.target="CarWin";
		CarForm.submit();
		CarForm.action="";
		CarForm.target="";
}
function funPrintStyle(){

		UrlStr="SendStyle.asp";
		newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
		myForm.action="SendStyle.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
//大宗郵件
function funMailList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印大宗郵件清冊的舉發單！");
	}else{
		UrlStr="MailSendList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailSendList",300,125,200,100,"no","no","no","no");
	}
}
//郵費清單
function funMailMoneyList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印郵費單的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList",300,160,350,200,"no","no","no","no");
	}
}
//逕舉
function funReportSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印逕舉移送清冊的舉發單！");
	}else{
		UrlStr="ReportSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停
function funStopSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印攔停移送清冊的舉發單！");
	}else{
		UrlStr="StopSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊
function funValidSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊
function funUselessSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//結案清冊
function funCaseCloseSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="CaseCloseSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"CaseCloseWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊
function funReturnSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊
function funStoreSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印寄存送達清冊的舉發單！");
	}else{
		UrlStr="funStoreSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin7",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊
function funGovSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//車籍查詢
function funchgCarDataList(){
	if (myForm.DB_Display.value==""){
		alert("請先查詢欲列印車籍清冊的舉發單！");
	}else{
		UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>&strDCISQL=<%=strDCISQL%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","no","yes","no");
	}
}
//影項車籍查詢
function funImgCarDataList(){
	if (myForm.DB_Display.value==""){
		alert("請先查詢數位影像車籍清冊的舉發單！");
	}else{
	<%if sys_City="嘉義縣" then%>
		UrlStr="DciImgCarDataList_CYS.asp?SQLstr=<%=strwhereToPrintCarData%>&strDCISQL=<%=strDCISQL%>";
	<%else%>
		UrlStr="DciImgCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>&strDCISQL=<%=strDCISQL%>";
	<%end if %>
		SCheight=screen.availHeight-50;
		SCWidth=screen.availWidth;

		window.open(UrlStr,"WebPage_img","left=0,top=0,location=0,width="+SCWidth+",height="+SCheight+",resizable=yes,scrollbars=yes,menubar=no")
	}
}

function funDCILogCarListDetail(){
	if (myForm.DB_Display.value==""){
		alert("請先查詢車籍清冊的舉發單！");
	}else{
		CarForm.action="CarAddressListDetail.asp";
		CarForm.target="inputWin";
		CarForm.submit();
		CarForm.action="";
		CarForm.target="";
	}
}

function funBillCaseDataDetail(){
	if (myForm.Sys_BatchNumber.value==""){
		alert("請先輸入作業批號！");
	}else{
		UrlStr="BillBaseCaseInDataDetail.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funStopCarListDetail(){
	if (myForm.DB_Display.value==""){
		alert("請先查詢車籍清冊的舉發單！");
	}else{
		CarForm.action="StopCarListDetail.asp";
		CarForm.target="inputWin";
		CarForm.submit();
		CarForm.action="";
		CarForm.target="";
	}
}
function funchgCkeckDataList(){
	if (myForm.Sys_BatchNumber.value==""){
		alert("請先輸入車籍查詢批號！");
	}else{
		UrlStr="CkeckOwnerDataList.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value;
		window.open(UrlStr,"WebPage_img","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes,menubar=no")
	}
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

</script>
<%conn.close%>