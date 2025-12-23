<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>各式清冊/舉發單列印</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
If isEmpty(request("DB_Display")) Then
	Sys_Now1=DateAdd("d",-2,date)&" "&hour(time)&":"&Minute(time)&":"&Second(time)
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
	End If 
	
	strSQL="select BillNo from BillMailHistory where Exists(select 'Y' from DCILog where ExchangeDate between TO_DATE('"&Sys_Now2&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&Sys_Now1&"','YYYY/MM/DD/HH24/MI/SS') and substr(batchnumber,1,1)='W' and RecordMemberID ="&Session("User_ID")&" and billsn=BillMailHistory.BillSN) and Exists(select 'Y' from Billbase where sn=BillMailHistory.BillSN and BillTypeID=2 and UseTool<>8 and RecordStateid=0) and maildate is null"

	chkbat=""

	set rschk=conn.execute(strSQL)
	while not rschk.eof
		If Not ifnull(chkbat) then chkbat=chkbat&"\n"
		chkbat=chkbat&rschk("BillNo")
		rschk.movenext
	wend
	rschk.close
	If not ifnull(chkbat) Then
		Response.write "<script>"
		Response.Write "alert('您下列單號尚未郵寄，請盡速確認！\n"&BillNo&"');"
		Response.write "</script>"
	End if
End if

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

'===================================================================
'strSQL="select * from UnitInfo where rownum=1"
'set rs=conn.execute(strSQL)
'If Not rs.eof Then
'	For i=0 to rs.Fields.count-1
'		If trim(rs.Fields.item(i).Name)="COUNTY" Then Exit For
'	Next
'	If i>rs.Fields.count-1 Then
'		strSQL="Alter Table UnitInfo ADD (COUNTY VarChar2(20))"
'		conn.execute(strSQL)
'	End if
'	For i=0 to rs.Fields.count-1
'		If trim(rs.Fields.item(i).Name)="UNITORDER" Then Exit For
'	Next
'	If i>rs.Fields.count-1 Then
'		strSQL="Alter Table UnitInfo ADD (UNITORDER VarChar2(6))"
'		conn.execute(strSQL)
'	End if
'
'	For i=0 to rs.Fields.count-1
'		If trim(rs.Fields.item(i).Name)="RECORDSTATEID" Then
'			strSQL="ALTER TABLE unitinfo RENAME COLUMN RecordStateid TO FiledStateid"
'			conn.execute(strSQL)
'			exit for
'		end if
'	Next
'End if
'rs.close
'===================================================================

RecordDate=split(gInitDT(date),"-")
if request("DB_Selt")="BatchSelt" Then
	strwhere_G8ML="" '苗栗無效清冊
	strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
	if UCase(request("Sys_BatchNumber"))<>"" then
		tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
		for i=0 to Ubound(tmp_BatchNumber)
			if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
			if i=0 then
				Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(tmp_BatchNumber(i))
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(tmp_BatchNumber(i))
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
		next
		strwhere=" and a.BatchNumber in('"&trim(Sys_BatchNumber)&"')"
		strwhere_G8ML=strwhere_G8ML&" and sn in (select billsn from dcilog where BatchNumber in('"&trim(Sys_BatchNumber)&"'))"
	end if

	if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.billsn in(select sn from billbase where billno between '"&trim(UCase(request("Sys_BillNo1")))&"' and '"&trim(UCase(request("Sys_BillNo2")))&"') and a.billno is not null"
		strwhere_G8ML=strwhere_G8ML&" and billno between '"&trim(UCase(request("Sys_BillNo1")))&"' and '"&trim(UCase(request("Sys_BillNo2")))&"'"
	elseif trim(request("Sys_BillNo1"))<>"" then
		strwhere=strwhere&" and a.billsn in(select sn from billbase where billno='"&trim(UCase(request("Sys_BillNo1")))&"') and a.billno is not null"
		strwhere_G8ML=strwhere_G8ML&" and billno='"&trim(UCase(request("Sys_BillNo1")))&"'"
	elseif trim(request("Sys_BillNo2"))<>"" then
		strwhere=strwhere&" and a.billsn in(select sn from billbase where billno='"&trim(UCase(request("Sys_BillNo2")))&"') and a.billno is not null"
		strwhere_G8ML=strwhere_G8ML&" and billno='"&trim(UCase(request("Sys_BillNo2")))&"'"
	end if
	
	if strwhere<>"" then
		strwhereToPrintCarData=strwhere
	else
		strwhereToPrintCarData=""
	end if

	if request("RecordDate")<>"" and request("RecordDate1")<>""then
		RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
		RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"
		if strwhere<>"" then
			if sys_City="台南市" then
				strwhere=strwhere&" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and a.RecordMemberID in (select MemberID from MemberData where UnitID='"&Session("Unit_ID")&"')"
			ElseIf sys_City="苗栗縣" Then
				strwhere=strwhere&" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and f.RecordMemberID <> 3552"

				strwhere_G8ML=strwhere_G8ML&" and RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID <> 3552"
			else
				strwhere=strwhere&" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and (f.RecordMemberID="&Session("User_ID")&" or a.RecordMemberID="&Session("User_ID")&")"
			end if
		else
			if sys_City="台南市" then
				strwhere=" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and a.RecordMemberID in (select MemberID from MemberData where UnitID='"&Session("Unit_ID")&"')"
			ElseIf sys_City="苗栗縣" Then
				strwhere=" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and f.RecordMemberID <> 3552"

				strwhere_G8ML=strwhere_G8ML&" and RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID <> 3552"
			else
				strwhere=" and f.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and (f.RecordMemberID="&Session("User_ID")&" or a.RecordMemberID="&Session("User_ID")&")"
			end if
		end If 
		
		If not ifnull(Request("Sys_Back")) Then
			strwhere=strwhere&" and a.exchangetypeid='"&trim(Request("Sys_Back"))&"'"	
		End if 

		If not ifnull(Request("Sys_BackTypeID")) Then
			If trim(Request("Sys_BackTypeID")) = "3" Then
				strwhere=strwhere&" and f.billtypeid='2' and f.UseTool<>8"
			else
				strwhere=strwhere&" and f.billtypeid='"&trim(Request("Sys_BackTypeID"))&"'"	
			End if 
			
		End if 
	end If 

	If sys_City="苗栗縣" Then

		If trim(Session("Ch_Name")) = "消防局入案" Then

			strwhere=strwhere&" and f.recordmemberid=3779"
		else

			strwhere=strwhere&" and f.recordmemberid<>3779"
		End if 		
	End if
	
end if
DB_Display=request("DB_Display")
if DB_Display="show" then
	if trim(strwhere)<>"" then
		'strwhereToPrintCarData=strwhere
'		if sys_City="屏東縣" then
'			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','n') "&strwhere&" and NVL(f.EquiPmentID,1)<>-1) or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','n') and a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') "&strwhere&" and NVL(f.EquiPmentID,1)<>-1) or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&" and NVL(f.EquiPmentID,1)<>-1)"
'		else
			If sys_City="基隆市" then
				KindType="('1','3','9','a','j','A','H','K','L','T')"

			elseIf sys_City="台中市" and session("User_ID")=5751 then
				KindType="('1','3','9','a','j','A','H','K','T')"

			elseIf sys_City="苗栗縣" then
				KindType="('1','3','9','a','j','A','H','K','T')"

			elseIf sys_City="高雄市" then
				KindType="('1','3','9','a','j','A','H','K','T','n')"

			else
				KindType="('1','3','9','a','j','A','H','K','T','n')"

			End if
			tempSQL="where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.billno=i.billno and a.CarNo=i.CarNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.BillTypeID='2' and a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&strwhere&" and NVL(f.EquiPmentID,1)<>-1"
'		end if
		'tempSQL=tempSQL&" and f.EquiPmentID<>-1"

		'if trim(request("PBillSN"))="" then '與dci上下查詢不同
		chk_MailNumKind=0
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			strSQL="select distinct a.BillSN from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4,BillCloseID from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL
			
			If sys_City<>"基隆市" and sys_City<>"苗栗縣" and sys_City<>"宜蘭縣" then

				strSQL=strSQL&" and a.DciReturnStatusID<>'n'"
			end if

			strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
			chk_MailNumKind=1
		else
			strSQL="select distinct a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i "&tempSQL

			If instr(request("Sys_BatchNumber"),"WT")>0 Then strSQL=strSQL&" and f.Note like '2%'"

			strSQL=strSQL&" order by f.RecordDate"
		end if

		set rssn=conn.execute(strSQL)
			BillSN="":tempBillSN=""
			while Not rssn.eof
				If trim(tempBillSN)<>trim(rssn("BillSN")) Then
					tempBillSN=trim(rssn("BillSN"))
					if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
					BillSN=BillSN&trim(rssn("BillSN"))
				end if
				rssn.movenext
			wend
			rssn.close
		'end if

'		if sys_City="屏東縣" then
'			tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','n') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','n') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607')) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
'		else
			'tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in "&KindType&" "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in "&KindType&" "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
'		end if

		If instr(request("Sys_BatchNumber"),"WT")>0 Then 
			strSQL="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.billno=i.billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&strwhere&" and f.Note like '1%'"

			set chksuess=conn.execute(strSQL)
			fileCloseOwner=cdbl(chksuess("cnt"))
			chksuess.close

			strSQL="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.billno=i.billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&strwhere&" and f.Note like '2%'"

			set chksuess=conn.execute(strSQL)
			fileCloseDriver=cdbl(chksuess("cnt"))
			chksuess.close
		end if
		
		strSQL="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f,(select distinct BillNo,CarNo,Rule4 from BillBaseDCIReturn where EXCHANGETYPEID='W') i where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.billno=i.billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and not (a.DciErrorCarData in "&KindType&" and i.Rule4<>'2607' and a.billtypeid='2') "&strwhere

		set chksuess=conn.execute(strSQL)
		filsuess=cdbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from DCILog a,Billbase f where a.BillSN=f.sn "&strwhere&" and ((ExchangeTypeID='E' and DCIReturnStatusID='n') or (ExchangeTypeID='W' and DCIReturnStatusID in ('S','d','e')) or (ExchangeTypeID='N' and DCIReturnStatusID='n'))"

		set chksuess=conn.execute(strSQL)
		filClose=cdbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and d.DCIRETURNSTATUS='-1' "&strwhere
		set chksuess=conn.execute(strSQL)
		fildel=cdbl(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from DCILog a,BillBase f where a.BillSN=f.SN "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=cdbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and f.RecordStateId=-1 and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=cdbl(Dbrs("cnt"))
		Dbrs.close
		
'		if sys_City="屏東縣" then
'			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','n') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and d.DCIRETURNSTATUS='1'"&strwhere
'		else
			strCnt="select count(*) as cnt from DCILog a,DCIReturnStatus d,BillBase f where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.BillSN=f.SN and a.DciErrorCarData in "&KindType&" and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and f.RecordStateId=0 and d.DCIRETURNSTATUS='1'"&strwhere
'		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt=cdbl(Dbrs("cnt"))
		Dbrs.close
		'filsuess=filsuess-errCatCnt
		tmpSQL=strwhere
		strSQL2=strwhere&" and NVL(f.EquiPmentID,1)<>-1"
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end If 
%>
<body>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr height="30">
		<td bgcolor="#1BF5FF"><span class="style3">各式清冊/舉發單列印</span>
		<% 
		If sys_City <> "高雄市" Then
		%>
		<img src="space.gif" width="60" height="1"> <strong>請勿升級 Internet Explorer 7 . 避免套印舉發單出現異常</strong></img>
		<%
		end if
		%>
		<br>
		<a href="../printoption.exe"><span class="pagetitle">清除印表機邊界</span></a>&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="../smsx.txt"><span class="pagetitle">安裝套印軟體</span></a>

		<a href="../下載說明.doc" target="_blank"><span class="pagetitle">(請先下載說明)</span></a>

		<a href="../IE7.reg.txt"><span class="pagetitle">移除ie7自動縮放</span></a>
		<%If sys_City = "宜蘭縣" or sys_City = "台東縣" Then%>
			<a href="../PrintLegal.exe"><span class="pagetitle">舉發單下載</span></a>

			<a href="printbak.html" target="_blank"><span class="pagetitle">舉發單背面列印</span></a>
		<%end if%>

		<img src="space.gif" width="12" height="1"><a href="DciCarErrorData.asp" target="_blank">查看逕舉無效原因</a>

		<img src="space.gif" width="12" height="1">
		<a href="PrityStyle.doc" target="_blank">舉發單紙張大小建立方式</a>

		<img src="space.gif" width="12" height="1">				
 		<a href="PrintError1.html" target="_blank">度量單位設定方式</a>		

		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							if sys_City="雲林縣" then
								nowdate=-2
							elseif sys_City="基隆市" then
								nowdate=-3
							else
								nowdate=-5
							end if

							strSQL="select Max(ExchangeDate) ExchangeDate,BatchNumber from DCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",nowdate, date)&" 00:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID in ('W','N','A') group by BatchNumber order by ExchangeDate DESC"

							if sys_City="苗栗縣" then
								strSQL="select Max(ExchangeDate) ExchangeDate,BillUit,BatchNumber from " & _
								"(select ExchangeDate,BatchNumber,(select (select (select UnitName from Unitinfo a where UnitID=b.UnitTypeID) from Unitinfo b where UnitID=BillBase.BillUnitID) from BillBase where SN=DciLog.BillSN) BillUit" & _
								" from DciLog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",nowdate, date)&" 00:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') and ExchangeTypeID in ('W','N','A'))" & _
								" group by BillUit,BatchNumber order by ExchangeDate DESC"
							end if
							
		
							set rs=conn.execute(strSQL)
							cut=0
							while Not rs.eof
								ExchangeDate=gInitDT(trim(rs("ExchangeDate")))
								if sys_City="苗栗縣" then
									response.write "<option value="""&trim(rs("BatchNumber"))&""">"
									response.write ExchangeDate& " - "&trim(rs("BillUit"))&"　"&trim(rs("BatchNumber"))
									response.write "</option>"
								else
									response.write "<option value="""&trim(rs("BatchNumber"))&""">"
									response.write ExchangeDate& " - "&cut&"　"&trim(rs("BatchNumber"))
									response.write "</option>"
								end if
								cut=cut+1
								rs.movenext
							wend
							rs.close
						%>
						</select>
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" size="95" onkeyup="funShowBillNo()">
						
						(<strong>多個批號同時處理</strong>，用,隔開。如：95A361,95A382）						
						<br>
						舉發單號
						<input name="Sys_BillNo1" id="Sys_BillNo1" type="text" class="btn1" value="<%=UCase(request("Sys_BillNo1"))%>" onkeyup="if(this.value.length>=9){myForm.Sys_BillNo2.focus();}" size="14" maxlength="9">
						~
						<input name="Sys_BillNo2" type="text" class="btn1" value="<%=UCase(request("Sys_BillNo2"))%>" size="13" maxlength="9"> ( 列印 <strong>單筆</strong> 或 特定範圍 舉發單才需填寫)
						<br>
						<%if sys_City<>"嘉義縣" then%>
							建檔日期
							<input name="RecordDate" type="text" value="<%=request("RecordDate")%>" size="11" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
							~
							<input name="RecordDate1" type="text" value="<%=request("RecordDate1")%>" size="10" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
							<%if sys_City="苗栗縣" then%>
								類別：
								<Select name="Sys_BackTypeID">
									<option value="">全部</option>
									<option value="3"<%If trim(Request("Sys_BackTypeID")) = "3" then Response.Write " selected" %>>逕舉(只有紅單)</option>
									<option value="2"<%If trim(Request("Sys_BackTypeID")) = "2" then Response.Write " selected" %>>逕舉(含手開單)</option>
									<option value="1"<%If trim(Request("Sys_BackTypeID")) = "1" then Response.Write " selected" %>>欄停</option>
								</select>
								項目：
								<Select name="Sys_Back">
									<option value="W">入案</option>
									<option value="N"<%If trim(Request("Sys_Back")) = "N" then Response.Write " selected" %>>收退件</option>
									<option value="A"<%If trim(Request("Sys_Back")) = "A" then Response.Write " selected" %>>車籍查詢</option>
								</select>
								<input type="submit" name="btnSelt" value="點收清冊" onclick="funAcceptDetialList();">
							<%end if%>
							(列印 入案移送清冊 / 大宗清冊 / 大宗掛號單 / 郵費單 可使用)
						<%else%>
							<img src="space.gif" width="60" height="1">舉發單以及各式清冊需要原案件建檔人才可列印
							<br>
							
							<input name="RecordDate" type="Hidden" value="" size="8" class="btn1">
							<input name="RecordDate1" type="Hidden" value="" size="8" class="btn1">
						<%end if%>
						<%if sys_City="苗栗縣" then
								strSQL="select distinct stationid,dcistationname from station order by stationid"
								set stat=conn.execute(strSQL)
								response.write "<br>監理處所：<Select Name=""Selt_MemberStation"">"
								response.write "<option value="""">"
								response.write "請選擇"
								response.write "</option>"
								while not stat.eof
									response.write "<option value="""&trim(stat("stationid"))&""""
									if trim(stat("stationid"))=trim(request("Selt_MemberStation")) then response.write " selected"
									response.write ">"
									response.write trim(stat("dcistationname"))
									response.write "</option>"
									stat.movenext
								wend
								response.write "</select>"
								stat.close
						else
							response.write "<input name=""Selt_MemberStation"" type=""Hidden"" value="""">"
						end if%>
						<br>
						<img src="space.gif" width="58" height="29">

						<%if sys_City<>"嘉義縣" then%>
							<input type="submit" name="btnSelt" value="查詢" onclick="funSelt('BatchSelt');">
						<%else%>
							<input type="button" name="btnSelt" value="查詢" onclick="funChiayiSelt('BatchSelt');">
						<%end if%>
						<input type="button" name="cancel" value="清除" onClick="location='AllBillPrint.asp'">
						

						<img src="space.gif" width="55" height="1"></img><strong>( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功(<%=filClose%>筆結案) , <%=errCatCnt%> 筆無效  ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-filsuess-fildel-deldata-errCatCnt%>筆未處理. )</strong>
						<br>
						<img src="space.gif" width="1" height="1"></img><font size="2" >列印舉發單/各式清冊前，請先輸入 批號 或是 舉發單號 進行 查詢</font><font color="red"><B><span id="upUnitTxt"><%=Request("upUnitName")%></span>&nbsp;&nbsp;<span id="showBillNoA"></span>&nbsp;&nbsp;<span id="showBillNoB"></span></B></font>
					</td>
				</tr>
			</table>
		</td>
	</tr>

	<tr>
		<td height="35" bgcolor="#1BF5FF" align="left">
		<img src="space.gif" width="8" height="1"></img>
		<%
			If instr(request("Sys_BatchNumber"),"WT")>0 Then
				Response.Write "<strong>"
				Response.Write fileCloseOwner&"筆車主已繳費，"
				Response.Write fileCloseDriver&"筆非車主已繳費。"
				Response.Write "</strong>"
			end if
		%>
		<img src="../Image/space.gif" width="40" height="10">
		<br><!-- titan測試用 -->
			<img src="space.gif" width="9" height="1"></img>
			<input type="button" name="btnprint" value="整批入案 郵寄日期/大宗條碼資料註記" onclick="funBillMailInfoMark()">

			<%if sys_City="嘉義縣" and Session("UnitLevelID") =1 then %>
				<img src="space.gif" width="8" height="1"></img>									
				<input type="button" name="cancel" value="舉發單代印工作" onClick="newWin('UpDateBillPrintJob.asp','inputWin',1200,550,50,10,'yes','yes','yes','no');"  style="width: 207px; height: 27px;">
			<%end if%>
			<%if sys_City="基隆市" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(0)"><br>

				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 14回執式舉發單 )" onclick="funBillNoPrint(26)">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 新版 )" onclick="funBillNoPrint(33)">
			<%elseif sys_City="台中市" and trim(Session("UnitLevelID"))="1" then%>
			<!--<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(18)">-->
				<br>
				<input name="chktelunit" type="radio" value="0469">直一電話　
				<input name="chktelunit" type="radio" value="0561">直三電話<br>
				<!--<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(19)"><br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 OKI 違規通知單" onclick="funBillNoPrintStyle(19)">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(回執聯104年)" onclick="funBillNoPrint(45)">-->
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(回執聯105年)" onclick="funBillNoPrint(46)">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(新版 回執聯105年)" onclick="funBillNoPrint(47)">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(新版107年)" onclick="funBillNoPrint(52)"><br>				
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 舉發單(107年套印數位相片) " onclick="funBillNoPrint(55)"><br>

			<%elseif sys_City="屏東縣" then%>
				<!--
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 17郵簡式舉發單 )" onclick="funBillNoPrint(16)">
				-->
				<br><img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 98 郵簡含送達證書 使用  ) " onclick="funBillNoPrint(79);">
				<br><br><img src="space.gif" width="8" height="1"></img>

				<input type="button" name="btnprint" value="列印 違規通知單( 2013 新版本  ) " onclick="funBillNoPrint(270);">
				<br><img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 含相片  ) " onclick="funBillNoPrint(179);">
				<br><img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 交通隊違規通知單(107年版) " onclick="funBillNoPrint(51);">
				<br><img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 各分局違規通知單(107年版) " onclick="funBillNoPrint(54);">
				<img src="space.gif" width="57" height="1"></img>

			<%elseif sys_City="花蓮縣" then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( A4 郵簡式舉發單 )" onclick="funBillNoPrint(1)">
				<img src="space.gif" width="57" height="1">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 郵簡式含違規影像 )" onclick="funBillNoPrint(30)">
				
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 停管郵簡式含違規影像 )" onclick="funBillNoPrint(63)">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 民眾檢舉用不含相片 )" onclick="funBillNoPrint(64)">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 民眾檢舉違規相片" onclick="PrintPicture_HuaLien();">

			<%elseif sys_City="台東縣" then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( A4 郵簡式舉發單 )" onclick="funBillNoPrint(15)">
				<img src="space.gif" width="7" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 郵簡含送達證書 )" onclick="funBillNoPrint(78)">
				<img src="space.gif" width="7" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 郵簡含大宗掛號碼 )" onclick="funBillNoPrint(25)">
				<img src="space.gif" width="7" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 空白套印 )" onclick="funBillNoPrint(34)">
			
			<%elseif sys_City="高港局" then %>
				<img src="space.gif" width="8" height="1"></img>									
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單 ,下聯相片  ) " onclick="funBillNoPrint(27);">
				<input type="button" name="btnprint" value="列印 【新版】違規通知單(  A4 雙色舉發單 ,下聯相片  ) " onclick="funBillNoPrint(271);">
			
			<%elseif sys_City="保二總隊三大隊二中隊" then '中科 %>
				<input type="button" name="btnprint" value="違規通知單(新版)" onclick="funBillNoPrint(59);"  style="width: 207px; height: 27px;">

			<%elseif sys_City="保二總隊三大隊一中隊" then '竹科 %>
				<input type="button" name="btnprint" value="違規通知單(新版)" onclick="funBillNoPrint(59);"  style="width: 207px; height: 27px;">

				<input type="button" name="btnprint" value="違規通知單(A4)" onclick="funBillNoPrint(60);"  style="width: 207px; height: 27px;">

			<%elseif sys_City="高雄市" and Session("UnitLevelID") =1 then %>
				<img src="space.gif" width="8" height="1"></img>									
				<input type="button" name="cancel" value="舉發單代印工作" onClick="newWin('UpDateBillPrintJob.asp','inputWin',1200,550,50,10,'yes','yes','yes','no');"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="陳情電話套印設定 " onclick="funBillNoTel();">
				<input type="button" name="btnprint" value="列印 違規通知單(三聯式)" onclick="funBillNoPrint(81);"  style="width: 207px; height: 27px;">
				<br>
				<input type="button" name="btnprint" value="違規通知單(新版)" onclick="funBillNoPrint(29);"  style="width: 207px; height: 27px;">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" name="btnprint" value="分局/分隊 違規通知單(新版)" onclick="funBillNoPrint(35);"  style="width: 207px; height: 27px;">
				<br>
				<input type="button" name="btnprint" value="違規通知單(新版不含照片)" onclick="funBillNoPrint(31);"  style="width: 207px; height: 27px;">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" name="btnprint" value="分局/分隊 違規通知單(新版不含照片)" onclick="funBillNoPrint(36);"  style="width: 207px; height: 27px;">
				<br>
				<%If trim(Session("Credit_ID"))="A000000000" or trim(Session("Credit_ID"))="12345" then %>
					<input type="button" name="btnprint" value="違規通知單(民眾檢舉網址)" onclick="funBillNoPrint(56);"  style="width: 207px; height: 27px;">

					<input type="button" name="btnprint" value="分局/分隊 違規通知單(民眾檢舉網址)" onclick="funBillNoPrint(57);"  style="width: 207px; height: 27px;">
				<%end if%>
				<!--
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 新版 舉發單封套 保防用" onclick="funLabelStyleKeelung_TaiChungCity();">
				-->
			<%elseif sys_City="苗栗縣" then %>
				<br>
				<img src="space.gif" width="8" height="1"></img>	
				停管違規事實補充說明：
				<input name="Sys_Rule4" type="text" class="btn1" size="30" value="">
				<input type="button" name="btnprint" value="更新" onclick="funUpdateRule4();">

				<br>
				<img src="space.gif" width="8" height="1"></img>									
				<input type="button" name="btnprint" value="舉發單(苗)" onclick="funBillNoPrint(37);"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="保防標籤(苗)" onclick="funLabelStyleLabel_miaoli();">
				<br>
				<img src="space.gif" width="8" height="1"></img>									
				<input type="button" name="btnprint" value="舉發單(苗)TEST" onclick="funBillNoPrint(41);"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單(108年版) " onclick="funBillNoPrint(61);">
				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="舉發單(108年區間速度版) " onclick="funBillNoPrint(37);">
			<%end if%>
			<%if sys_City="台中市" then%>
				<input type="button" name="cancel" value="舉發單代印工作" onClick="newWin('UpDateBillPrintJob.asp','inputWin',1200,550,50,10,'yes','yes','yes','no');"  style="width: 207px; height: 27px;">
			<%end if%>
			<%if sys_City="澎湖縣" then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(2);">
				<img src="space.gif" width="57" height="1">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 1030501  ) " onclick="funBillNoPrint(39);">
			<%elseif sys_City="金門縣" then %>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 1030501  ) " onclick="funBillNoPrint(2);">
				<img src="space.gif" width="57" height="1">				
				<input type="button" name="btnprint" value="列印 違規通知單(108年版) " onclick="funBillNoPrint(62);">
				<img src="space.gif" width="57" height="1">				
				<input type="button" name="btnprint" value="列印 違規通知單(108年版無照片) " onclick="funBillNoPrint(65);">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 1060315  ) " onclick="funBillNoPrint(48);">
			<%elseif sys_City="嘉義市" then %>
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(21);">
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 103年版  ) " onclick="funBillNoPrint(42);">
				<img src="space.gif" width="57" height="1">
			<%elseif sys_City="高雄縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="員警印章套印設定 " onclick="funSealBillNo();">
				<img src="space.gif" width="8" height="1"></img>
				<!--<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單 ,下聯空白  ) " onclick="funBillNoKaoHsiungPrint();">-->
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單 ,下聯空白  ) " onclick="funBillNoPrint(17);">
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 存根聯  ) " onclick="funBillNoPrint(20);">
				<input type="button" name="btnprint" value="列印 違規通知單( 停車入案 使用  ) " onclick="funBillNoPrint(76);">
				<input type="button" name="btnprint" value="補印 違規通知單(三聯式)" onclick="funLegalPrintMend_KaoHsiungMend();"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="57" height="1">								
			<%elseif sys_City="彰化縣" or sys_City="台南縣" then%>
				<br>
				<!--
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(7);">
				-->
				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 新版  ) " onclick="funBillNoPrint(38);">


				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="列印 (交通隊)違規通知單(107年版) " onclick="funBillNoPrint(53);">

				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="列印 (分局)違規通知單(107年版) " onclick="funBillNoPrint(58);">
			
				<br>
			<%elseif sys_City="嘉義縣" then%>
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<!--<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(10);">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 新式違規通知單(A4) " onclick="funBillNoPrint(32);">
				<img src="space.gif" width="8" height="1"></img>-->
				<input type="button" name="btnprint" value="列印 103年違規通知單(A4) " onclick="funBillNoPrint(40);">
				<img src="space.gif" width="57" height="1">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="舉發單(含數位影像)" onclick="funBillNoPrint(43);"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="舉發單(不含數位影像)" onclick="funBillNoPrint(44);"  style="width: 207px; height: 27px;">
			<%elseif sys_City="雲林縣" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(9);">
				<img src="space.gif" width="57" height="1"><br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 郵簡含送達證書  ) " onclick="funBillNoPrint(77);">
				<img src="space.gif" width="57" height="1">				
			<%elseif sys_City="台南市" then%>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單(  A4 雙色舉發單  ) " onclick="funBillNoPrint(14);">
				<img src="space.gif" width="57" height="1">
			<%elseif sys_City="台中市" then%>
				<br>	
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單存根聯" onclick="funBillNoPrint(12);">
				<img src="space.gif" width="57" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單存根聯(106年)" onclick="funBillNoPrint(50)"><br>
			<%end if%>
			<%if sys_City="花蓮縣" then %>		
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="花蓮縣分局 列印違規通知單（ 點陣式 ）" onclick="funBillNoPrint(8)">
			<%elseif sys_City="宜蘭縣" then%>
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( 點陣式 )" onclick="funBillNoPrint(13)">
				<br>
<!--				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legeal )" onclick="funBillNoPrint(22)">
				<br>-->
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="列印 違規通知單( Legeal 新版 )" onclick="funBillNoPrint(23)">
				<br>
				<img src="space.gif" width="8" height="1"></img>
				<input type="button" name="btnprint" value="測試 違規通知單( Legeal 新版 )" onclick="funBillNoPrint(28)">

			<%elseif sys_City="連江縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 點陣式 )" onclick="funBillNoPrint(5)">
			<%elseif sys_City="南投縣" then%>
				<br>
	<img src="space.gif" width="8" height="1">郵寄地址選項規則: <br>
	<img src="space.gif" width="8" height="1">舉發單 &nbsp;&nbsp;&nbsp;<b>含</b>送達證書: 
					<b>第一次郵寄</b> (通-> 戶-> 車)&nbsp; &nbsp;
					<b>第二次郵寄</b>  (戶-> 車) <br>
	<img src="space.gif" width="8" height="1">舉發單<b>不含</b>送達證書: 
					<b>第一次郵寄</b> (通-> 車)						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<b>第二次郵寄</b> (戶-> 車) 
			<br>
<img src="space.gif" width="8" height="1">*需有進行車籍查詢 , 單退註記 後才有戶籍資料 <br>
<img src="space.gif" width="8" height="1">*交寄大宗函件請選擇有無含送達證書<br>

				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 點陣式 )" onclick="funBillNoPrint(6)">
				<input type="button" name="btnprint" value="列印 違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(80)">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="數位影像違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(82)">

				<br><img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="新式違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(83)">
<br>
				<br>

				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="測試 102 數位影像違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(101)">
<!--
				<br><img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="測試 102 新式違規通知單( Legal 8.5 X 14郵簡式舉發單 )" onclick="funBillNoPrint(102)">
				
				<br><img src="space.gif" width="18" height="1">
				<input type="button" name="btnprint" value="測試 新式違規通知單( Legal 8.5 X 14郵簡式舉發單 )(含相片)" onclick="funBillNoPrint(98)">
				-->
			<%elseif sys_City="台中縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( 點陣式 )" onclick="funBillNoPrint(11)">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 違規通知單( Legal )" onclick="funBillNoPrint(24)">
			<%end if%>
			<%if sys_City="基隆市" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 舊版 舉發單封套 保防用( Legal )" onclick="funLabelFormat();"><br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="新版 舉發單封套 (回執聯)" onclick="funLabelFormat_New();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="新版 舉發單封套 (送達書)" onclick="funLabelFormat_Deliver();">
				<input type="button" name="btnprint" value="新版 舉發單封套 (更正通知)" onclick="funLabelFormat_Update();">
				<input type="button" name="btnprint" value="新版 舉發單封套 (存根聯)" onclick="funLabelFormat_act();">
			<%elseif sys_City="台中市" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 新版 舉發單封套 保防用" onclick="funLabelStyleKeelung_TaiChungCity();">
			<%elseif sys_City="高雄縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 新版 舉發單封套 保防用" onclick="funLabelStyleKeelung_KaoHsiung();">
			<%elseif sys_City="高港局" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="送達證書" onclick="funLabelStyleKeelung_KaoHsiungHarBor();">

			<%elseif sys_City="保二總隊四大隊二中隊" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="送達證書" onclick="funLabelStyleKeelung_KaoHsiungHarBor();">

			<%elseif sys_City="高雄市" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 新版 舉發單封套 保防用" onclick="funLabelStyleKeelung_KaoHsiung();">
				<img src="space.gif" width="8" height="1">

				<input type="button" name="btnprint" value="補印 違規通知單(三聯式)" onclick="funLegalPrintMend_KaoHsiungHarBor();"  style="width: 207px; height: 27px;">

				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 舉發單封套(含送達證書)" onclick="funPasserUrgeHuaLien_DeliverListLabel();">

			<%elseif sys_City="苗栗縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="送達證書(苗)" onclick="funLabelStyleKeelung_miaoli();">

			<%elseif sys_City="南投縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 舉發單信封黏貼標籤" onclick="label_Style_Keelung_NanTou();">
			<%elseif sys_City="彰化縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 舉發單信封黏貼標籤" onclick="label_Style_Keelung_CHCG();">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="保防標籤套印（A4）" onclick="funLabelStyle();">
			<%elseif sys_City="台南市" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value=" 保防標籤套印（A4）" onclick="label_Style_TaiNaNCity();">
			<%elseif sys_City<>"保二總隊四大隊二中隊" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="保防標籤套印（A4）" onclick="funLabelStyle();">
			<%end if%>
			<br>
			<%if sys_City="南投縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillNonTouSendLegal();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillNonTouNewSendLegal();">
			<%elseif sys_City="台中縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillTaiChungSendLegal();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 台中市送達證書 ( 橫式 ) " onclick="funBillTaiChungCitySendLegal();">
			<%elseif sys_City="彰化縣" or sys_City="嘉義縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 直式 ) " onclick="funBillCHCGLegal();">
			<%elseif sys_City="花蓮縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillHuaLienSendLegal();">
			<%elseif sys_City="嘉義市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillChiayiCitySendLegal();">
			<%elseif sys_City="台中市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 舉發單封套(含送達證書)" onclick="funPasserUrgetaichung_DeliverListLabel();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="套印 舉發單封套(含送達證書)" onclick="funPasserUrgetaichung_Deliver_chromat();">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillTaiChungCitySendLegal();">
			<%elseif sys_City="高雄縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 ( 橫式 ) " onclick="funBillKaoHsiungSendLegal();">
			<%elseif sys_City="高雄市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit4234" value="逕舉手開單移送清冊 13.6 x 11 " onclick="funReportSendList_HL()"  style="width: 207px; height: 27px;">
				<input type="button" name="Submit4234" value="逕舉手開單移送清冊 A4" onclick="funReportSendList()"  style="width: 207px; height: 27px;">

			<%elseif sys_City="苗栗縣" then%>

				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit4234" value="逕舉手開單移送清冊 A4" onclick="funReportSendList()"  style="width: 207px; height: 27px;">

			<%elseif sys_City="高港局" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit4234" value="逕舉手開單移送清冊 A4" onclick="funReportSendList()"  style="width: 207px; height: 27px;">
			<%elseif sys_City="保二總隊四大隊二中隊" then%>

				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()"  style="width: 207px; height: 27px;">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="Submit4234" value="逕舉手開單移送清冊 A4" onclick="funReportSendList()"  style="width: 207px; height: 27px;">
			
			<%elseif sys_City="台東縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書套印 ( 橫式 )  " onclick="funBillTaiTungSendLegal();">				
				<input type="button" name="btnprint" value="列印 送達證書一式三份 ( 橫式 )  " onclick="funBillSendB5H();">
			<%end if%>
			
			<%if instr(sys_City,"高雄市")=0 and instr(sys_City,"高港局")=0 and instr(sys_City,"保二總隊四大隊二中隊")=0 then%>
				
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="列印 送達證書 (  B5  )  " onclick="funBillSendB5();">
				<input type="button" name="btnprint" value="列印 送達證書 (B5直式)  " onclick="funBillSendB_A4();">
				<input type="button" name="btnprint" value="列印 送達證書 (  A4  )  " onclick="funBillSendLegal();">
			<% end if %>
 
			
			<%if sys_City="彰化縣" or sys_City="南投縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="掛號郵件收回執 " onclick="funFastPostReceive();">
			<%elseif sys_City="花蓮縣" then%>
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯" onclick="funFastPostReceive_HuaLien();">
			<%end if%>
			<%if sys_City="彰化縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯A4(中一刀)" onclick="funFastPostReceiveA4();">
			<%end if%>
			<%if sys_City="南投縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯（新版）" onclick="funFastPostReceive_new();">
			<%elseif sys_City="台中縣" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯(郵局格式舊版)" onclick="funFastPostReceive_TaiChung();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯(郵局格式新版)" onclick="funFastPostReceive_TaiChung2();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯（新版）" onclick="funFastPostReceive_new();">
			<%elseif sys_City="台中市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="臺中縣回執聯" onclick="funFastPostReceive_new();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯(郵局提供格式)" onclick="funFastPostReceive_TaiChungCity();">
				<input type="button" name="btnprint" value="新回執聯(郵局提供格式)" onclick="funFastPostReceive_TaiChungCity2();">
			<%elseif sys_City="基隆市" then%>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回執聯(郵局提供格式)" onclick="funFastPostReceive_TaiChungCity();">
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="回收執(A4，一頁三張)" onclick="funFastPostReceive_Keelung();">
			<%end if%>
			
			
		<%If sys_City="高雄市" then%>
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit4335" value="攔停移送清冊 13.6x11" onclick="funStopSendList_A3_KSC()"  style="width: 207px; height: 27px;">		

			<!-- 
			<input type="button" name="Submit4335" value="拖吊已結移送清冊 A3" onclick="funMailNotBackList_TakeCar()"  style="width: 207px; height: 27px;">				
			-->
			<input type="button" name="Submit4335" value="攔停移送清冊 A4" onclick="funStopSendList_KSC()" style="width: 207px; height: 27px;">
				<br>
				<img src="space.gif" width="8" height="1">
				<input type="button" name="btnprint" value="送達證書 Letter(分局分隊套印)" onclick="funBillKaoHsiungCitySendLegal();"  style="width: 207px; height: 27px;">
				<input type="button" name="btnprint" value="列印 送達證書 A4 " onclick="funBillSendLegal();"  style="width: 207px; height: 27px;"><br>
				<img src="space.gif" width="8" height="1">
				<a href="A5DeliverSample.jpg" target="_blank"><span class="pagetitle">送達證書 Letter(範例)</span></a>
		
		<%elseif sys_City="苗栗縣" then%>
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit4335" value="攔停移送清冊 A4(所有監理站)" onclick="funStopSendList_KSC()" style="width: 207px; height: 27px;">
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit4335" value="攔停移送清冊 A4(苗栗監理站)" onclick="funStopSendList_ML()" style="width: 207px; height: 27px;">
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()"  style="width: 207px; height: 27px;">
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit3f32" value="交寄大宗函件(苗栗縣)" onclick="funMailListML2()"  style="width: 207px; height: 27px;">
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">

		<%elseif sys_City="高港局" then%>
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit4335" value="攔停移送清冊 A4" onclick="funStopSendList_KSC()" style="width: 207px; height: 27px;">
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="btnprint" value="列印 送達證書 A4 " onclick="funBillSendLegal();"  style="width: 207px; height: 27px;"><br>

		<%elseif sys_City="保二總隊四大隊二中隊" then%>
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="Submit4335" value="攔停移送清冊 A4" onclick="funStopSendList_KSC()" style="width: 207px; height: 27px;">
			<br>
			<img src="space.gif" width="8" height="1">
			<input type="button" name="btnprint" value="列印 送達證書 A4 " onclick="funBillSendLegal();"  style="width: 207px; height: 27px;"><br>

		<%end if%>
			
			
			
			<hr>
			
			<%if sys_City="台南市" then %>
				&nbsp;&nbsp;&nbsp;&nbsp; * <b>交寄大宗函件 </b> 如有需要顯示 <b>郵寄日期</b> ，請先輸入批號查詢後，由下方設定 該批資料 郵寄日期
				<br>			
				<br>	
			<%end if%>			
			<!--<span class="style3">
			DCI檔案名稱
			<input name="textfield42324" type="text" value="" size="14" maxlength="13">
			</span>-->
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市"  then '花蓮專用A3版%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList_HL()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊" onclick="funReportSendList_HL()">
		<%elseif sys_City <> "高雄市" and sys_City <> "保二總隊四大隊二中隊" and sys_City <> "苗栗縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234222" value="車籍資料" onclick="funchgCarDataList()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊" onclick="funReportSendList()">
		<%else%>
			<img src="space.gif" width="180" height="1">
		<%end if%>
		
		<span class="style3"><img src="space.gif" width="15" height="8"></span>
		
		<%If sys_City<>"高雄市" and sys_City<>"高港局" and sys_City <> "保二總隊四大隊二中隊" and sys_City <> "苗栗縣" then%>
			<input type="button" name="Submit42342" value="大宗掛號清冊" onclick="funMailList()">
			<img src="space.gif" width="8" height="8">
		<%else%>			
			<img src="space.gif" width="70" height="8">
		<%end if%>
		
		
		<input type="button" name="Submit488423" value="退件清冊_寄存( 全 部 )" onclick="funReturnSendList_Store_All()">
		
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="12" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊( 全 部 )" onclick="funStoreSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="15" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊( 全 部 )" onclick="funStoreSendList()">
		<%end if%>
		<br>
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then '花蓮專用A3版%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList_HL()">
		<%elseif sys_City="苗栗縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊(分單位)" onclick="funValidSendList_ML()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊" onclick="funValidSendList()">
		<%end if%>
			<%'if trim(request("Sys_ExchangeTypeID"))="W" then '入案%>
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4335" value="攔停移送清冊" onclick="funStopSendList_HL()">			
		<%elseif sys_City="高雄市" or sys_City="高港局" then%>
			<span class="style3"><img src="space.gif" width="150" height="8"></span>
		<%elseif sys_City="保二總隊四大隊二中隊" then%>
			<span class="style3"><img src="space.gif" width="5" height="8"></span>
		<%elseif sys_City="苗栗縣" then%>
			<span class="style3"><img src="space.gif" width="150" height="8"></span>

		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>			
			<input type="button" name="Submit4335" value="攔停移送清冊" onclick="funStopSendList()">
		<%end if%>
		
		<%If sys_City="保二總隊四大隊二中隊" Then%>
			<span class="style3"><img src="space.gif" width="8" height="8"></span>
			<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">
			<span class="style3"><img src="space.gif" width="79" height="8"></span>
		<%elseIf sys_City<>"高雄市" and sys_City<>"苗栗縣" then%>
			<span class="style3"><img src="space.gif" width="8" height="8"></span>
			<input type="button" name="Submit423423" value="郵費單" onclick="funMailMoneyList()">
			<span class="style3"><img src="space.gif" width="79" height="8"></span>
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			
		<%end if%>
			
			<input type="button" name="Submit488423" value="退件清冊_寄存(已結案)" onclick="funReturnSendList_Store_Close()">

			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(已結案)" onclick="funStoreSendList_Close()">
		<%if sys_City="苗栗縣X" Then '先不要%>
		<br>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit47335" value="有效清冊(分單位) 攔停案件用" onclick="funValidSendList_ML_Stop()">
		<%End If %>
		<br>
		<%if sys_City="苗栗縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="無效清冊(刪除)" onclick="funUselessSendList_ML()">			
		<%End If %>
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="無效清冊" onclick="funUselessSendList_HL()">			
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="無效清冊<%
			if sys_City="苗栗縣" Then
				response.write "(失敗)"
			End If 
			%>" onclick="funUselessSendList()">
		<%end if%>
			<%if sys_City<>"高雄市" then %>
				<span class="style3"><img src="space.gif" width="10" height="8"></span>
				<input type="button" name="Submit4234" value="郵寄未退回清冊" onclick="funMailNotBackList()" style="width: 135px; height: 27px;">
			<%else%>
				<span class="style3"><img src="space.gif" width="150" height="8"></span>
			<%end if%>
			
			
			<!-- <span class="style3"><img src="space.gif" width="163" height="8"></span> -->
			<%if instr(sys_City,"高雄市")=0 and instr(sys_City,"高港局")=0 and instr(sys_City,"保二總隊四大隊二中隊")=0 and instr(sys_City,"苗栗縣")=0 then%>
			     <span class="style3"><img src="space.gif" width="10" height="8"></span>
				<input type="button" name="Submit3f32" value="交寄大宗函件" onclick="funMailList2()">
				<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<%else%>
				<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<%end if %>
			
			<input type="button" name="Submit4233" value="退件清冊_寄存(未結案)" onclick="funReturnSendList_Store()">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4233" value="寄存送達清冊(未結案)" onclick="funStoreSendList_UnClose()">
		<br>
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="結案清冊" onclick="funCaseCloseSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit43635" value="結案清冊" onclick="funCaseCloseSendList()">
		<%end if%>
		<%if sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4234" value="逕舉移送清冊_A4" onclick="funReportSendList()" style="width: 135px; height: 27px;">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
		<%else%>
			
			<span class="style3"><img src="space.gif" width="12" height="8"></span>
		<%end if%>
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList_HL()">
		<%else%>
			<input type="button" name="Submit488423" value="收受清冊" onclick="funGetSendList()">
		<%end if%>
		<%if sys_City="高雄市" or sys_City="高港局" or sys_City="保二總隊四大隊二中隊" then%>
			
		<%else%>
			<span class="style3"><img src="space.gif" width="145" height="8"></span>
		<% end if %>
		
			<span class="style3"><img src="space.gif" width="55" height="8"></span>
			<input type="button" name="Submit4233" value="退件清冊_公示( 全 部 )" onclick="funReturnSendList_Gov_All()">
		
		<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<span class="style3"><img src="space.gif" width="33" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊( 全 部 )" onclick="funGovSendList_HL()">
		<%else%>
			<span class="style3"><img src="space.gif" width="15" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊( 全 部 )" onclick="funGovSendList()">
		<%end if%>
		<br>
			<span class="style3"><img src="space.gif" width="8" height="8"></span>
			<input type="button" name="Submit4232" value="不郵寄清冊" onclick="funNotMailList()" style="width: 90px; height: 27px;">
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<%if sys_City="嘉義縣" or sys_City="嘉義市" then%>
				<span class="style3"><img src="space.gif" width="10" height="8"></span>
				<input type="button" name="Submit4234" value="攔停移送清冊_A4" onclick="funStopSendList()" style="width: 135px; height: 27px;">
				<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<%else%>
				<span class="style3"><img src="space.gif" width="1" height="8"></span>
			<%end if%>
			<input type="button" name="Submit4232" value="公告清冊" onclick="funOpenGovList()">
			
			<%if sys_City="高雄市" or sys_City="高港局" or sys_City="保二總隊四大隊二中隊" then%>
				<span class="style3"><img src="space.gif" width="55" height="8"></span>
			<%else%>
				<span class="style3"><img src="space.gif" width="205" height="8"></span>
			<% end if %>			
			<input type="button" name="Submit488423" value="退件清冊_公示(已結案)" onclick="funReturnSendList_Gov_Close()">
			
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(已結案)" onclick="funGovSendList_Close()">
		<br>
		
		<%if sys_City="高雄市" or sys_City="高港局" or sys_City="保二總隊四大隊二中隊" then%>
			<span class="style3"><img src="space.gif" width="274" height="8"></span>
		<%else%>
			<span class="style3"><img src="space.gif" width="424" height="8"></span>
		<%end if%>
			<input type="button" name="Submit488423" value="退件清冊_公示(未結案)" onclick="funReturnSendList_Gov()">		
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(未結案)" onclick="funGovSendList_UnClose()">
		<br>
		<%if sys_City="南投縣" then%>
			<span class="style3"><img src="space.gif" width="10" height="8"></span>
			<input type="button" name="Submit4232" value="公示送達清冊(全部)_不分頁" onclick="funGovSendList_NoPage()">
		<%end if%>
		<%if sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<input type="button" name="Submit4233" value="退件清冊_寄存(未結案)_A4" onclick="funReturnSendList_Store_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="退件清冊_寄存(已結案)_A4" onclick="funReturnSendList_Store_Close_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="退件清冊_公示(未結案)_A4" onclick="funReturnSendList_Gov_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="退件清冊_公示(已結案)_A4" onclick="funReturnSendList_Gov_Close_A4()" style="width: 210px; height: 27px;">
			<br>
		<%end if%>
		<%if sys_City="嘉義縣" or sys_City="嘉義市" then%>
			<input type="button" name="Submit4233" value="寄存送達清冊(未結案)_A4" onclick="funStoreSendList_UnClose_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="寄存送達清冊(已結案)_A4" onclick="funStoreSendList_Close_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="公示送達清冊(未結案)_A4" onclick="funGovSendList_UnClose_A4()" style="width: 210px; height: 27px;">
			<input type="button" name="Submit488423" value="公示送達清冊(已結案)_A4" onclick="funGovSendList_Close_A4()" style="width: 210px; height: 27px;">
		<br>
		<%end if%>
		<br>
		<%if sys_City="台中市" then%>
			<input type="button" name="Submit488423" value="拖吊未繳費已領單清冊" onclick="funReporTrailertSend();" style="width: 210px; height: 27px;">
		<%end if%>
		<HR>
		
		<%if sys_City<>"保二總隊四大隊二中隊" then%>
			本批資料<b>第一次郵寄日期</b>
			&nbsp;&nbsp;&nbsp;<input name="Sys_BillBaseMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BillBaseMailDate');">
			&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSys_MailDate();"<%
			if Instr(request("Sys_BatchNumber"),"N")>0 then Response.Write "disabled"%>>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			本批資料<b>第二次郵寄日期</b>
			<input name="Sys_StoreAndSendMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendMailDate');">
			&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendMailDate();"<%
			if Instr(request("Sys_BatchNumber"),"N")=0 then Response.Write "disabled"%>>
			<br>
			本批資料發文監理站日期
			<img src="space.gif" width="15" height="1"> 
			<input name="Sys_SendOpenGovDocToStationDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendOpenGovDocToStationDate');">
			&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSendOpenGovDocToStationDate();">
		<%end if%>

		<%if sys_City="基隆市" then%>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			本批資料送達郵寄日期
			<input name="Sys_StoreAndSendFinalMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendFinalMailDate');">
			&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendFinalMailDate();">
		<%end if%>
		<br>
		<%if sys_City="花蓮縣" or sys_City="台中縣" then %>
			<hr>
			<b>寄存送達期滿</b>註記日期&nbsp;&nbsp;&nbsp;&nbsp;
			<input name="Sys_BillStoreFinish1" type="text" class="btn1" size="10" maxlength="11">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BillStoreFinish1');">
			～
			<input name="Sys_BillStoreFinish2" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BillStoreFinish2');">
			寄存期滿文號</b>
			&nbsp;&nbsp;&nbsp;<input name="Sys_StoreAndSendNumber" type="text" class="btn1" size="10">
			&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funStoreAndSendNumber();">
			<br>
			&nbsp;&nbsp;<input type="button" name="btnOK" value="清冊(區分監理站)" onclick="funBillStoreFinish();">
			&nbsp;&nbsp;<input type="button" name="btnOK2" value="清冊(不區分監理站)" onclick="funBillStoreFinish2();">
		<%end if%>
		<br><br>
	</td>
  </tr>
  <tr>
    <td><p align="center">&nbsp;</p>    </td></tr>
</table>
<input type="button" name="Submit4232" value="違規舉發單 / 清冊 列印設定說明" onclick="funPrintDetail()"> 
  各式清冊依據縣市需求分為 A4  或 A3 格式 或 13.6x11 </br> 
<br>
	
	<font size="5"> 
	列印 <b>各式清冊 <br>
	超出頁面 <img src="space.gif" width="40" height="1"></b> 請確認 檔案 --> 列印格式--> 紙張設定<font size="3"> 
	(請依據縣市需求選擇A4或A3或13.6x11)</font><br>
	<img src="space.gif" width="450" height="1">上下左右邊界請設定是否皆為 0mm 或是 5.08mm <br>
	<b>頁尾出現網址 </b> 請確認 檔案 --> 列印格式--> 頁首頁尾皆為空白 
    </font>
	<br />
	<%if sys_City="雲林縣" then %>
	<a href="clear.exe" target="_blank"><font  size="3">清除IE程式(請按右鍵另存目標)</font></a>
	<%End if%>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="upUnitName" value="">
<input type="Hidden" name="hd_PrintSum" value="0">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="PBillSN" value="<%=BillSN%>">
<input type="Hidden" name="printStyle" value="">
<input type="Hidden" name="Sys_MailDate" value="">
<input type="Hidden" name="Sys_JudeAgentSex" value="">
<input type="Hidden" name="Sys_Print" value="">
<input type="Hidden" name="Sys_CityKind" value="0">
<input type="Hidden" name="Sys_strSQL" value="<%=strSQL2%>">
<input type="Hidden" name="billprintuseimage" value="">
<input type="Hidden" name="Sys_AllPrintSQL" value="<%=tmpSQL%>">
<input type="Hidden" name="hd_BillJobName" value="">
<input type="Hidden" name="hd_MainChName" value="">
<input type="Hidden" name="Sys_SendKind" value="">
<input type="Hidden" name="Sys_LabelKind" value="">
<input type="Hidden" name="Sys_LabelUpdate" value="">
<input type="Hidden" name="Sys_UnitLabelKind" value="">
<input type="Hidden" name="Sys_BillPrintUnitTel" value="">
<input type="Hidden" name="Sys_label_Stytle_location" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var winopen;
funShowBillNo();
var sys_City='<%=sys_City%>';

/*function escKeyPress(){
	var btn = document.getElementById('Sys_BillNo1');
	var evt = document.createEvent('KeyboardEvent');

}*/

function funUpdateRule4(){
	if(myForm.Sys_BatchNumber.value!=''&&myForm.Sys_Rule4.value!=''){

		runServerScript("UpdateRule4_miaoli.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_Rule4="+myForm.Sys_Rule4.value);

	}
}

function funShowBillNo(){
	if(myForm.Sys_BatchNumber.value.length>=5){
		runServerScript("chkShowBillNo.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value);
	}
}

function funAcceptDetialList(){
	if (myForm.RecordDate1.value==""){
		alert("請先輸入建檔日期！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="AcceptDetailList.asp";
		myForm.action=UrlStr;
		myForm.target="funAcceptDetialList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funSendOpenGovDocToStationDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_SendOpenGovDocToStationDate.value!=''){
			//runServerScript("SendToStationDate.asp?SendOpenDate="+myForm.Sys_SendOpenGovDocToStationDate.value+"&BillSn="+myForm.PBillSN.value);
			var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			xmlhttp.Open("post","SendToStationDate.asp",false);	
			xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");			
			xmlhttp.send("SendOpenDate="+myForm.Sys_SendOpenGovDocToStationDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
			alert("儲存完成!!");
		}
	}
}

function funStoreAndSendFinalMailDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_StoreAndSendFinalMailDate.value!=''){
			var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			xmlhttp.Open("post","StoreAndSendFinalMailDate.asp",false);	
			xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");			
			xmlhttp.send("StoreAndSendFinalMailDate="+myForm.Sys_StoreAndSendFinalMailDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
			alert("儲存完成!!");
		}
	}
}

function funSys_MailDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_BillBaseMailDate.value!=''){
			
			var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			xmlhttp.Open("post","BillBaseMailDate.asp",false);	
			xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");
			xmlhttp.send("MailDate="+myForm.Sys_BillBaseMailDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
			alert("儲存完成!!");
		}
	}
}

function fnBatchNumber(){
	funShowBillNo();
	if(sys_City=="苗栗縣"){
		if(myForm.Sys_BatchNumber.value!=''){myForm.Sys_BatchNumber.value=myForm.Sys_BatchNumber.value+',';}
		myForm.Sys_BatchNumber.value=myForm.Sys_BatchNumber.value+myForm.Selt_BatchNumber.value;
	}else{
		myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
	}
}

function funStoreAndSendMailDate(){
	if (myForm.DB_Display.value!=""){
		if (myForm.Sys_StoreAndSendMailDate.value!=''){
			//runServerScript("StoreAndSendMailDate.asp?StoreAndSendMailDate="+myForm.Sys_StoreAndSendMailDate.value+"&BillSn="+myForm.PBillSN.value);
			var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
			xmlhttp.Open("post","StoreAndSendMailDate.asp",false);	
			xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded;");
			xmlhttp.send("StoreAndSendMailDate="+myForm.Sys_StoreAndSendMailDate.value+"&Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
			//alert(xmlhttp.responsetext);
			alert("儲存完成!!");
		}
	}
}

function funChiayiSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==""&&myForm.Sys_BillNo1.value==""&&myForm.Sys_BillNo2.value==""&&myForm.RecordDate.value==""&&myForm.RecordDate1.value==""){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(error==0){
			runServerScript("chkAllBillPrint.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value);
		}
	}
}

function funSelt(DBKind){
	var error=0;
	if(DBKind=='BatchSelt'){
		if(myForm.Sys_BatchNumber.value==""&&myForm.Sys_BillNo1.value==""&&myForm.Sys_BillNo2.value==""&&myForm.RecordDate.value==""&&myForm.RecordDate1.value==""){
			error=1;
			alert("必須有填詢條件!!");
		}
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=1;
				alert("建檔日期輸入不正確!!");
			}
		}
		if(error==0){
			myForm.hd_PrintSum.value="0";
			myForm.PBillSN.value="";
			myForm.DB_Move.value="";
			myForm.DB_Selt.value=DBKind;
			myForm.DB_Display.value='show';
			myForm.submit();
		}
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}
function newWin2(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	winopen.focus();
	return win;
}
function funDataDetail(SN){
	UrlStr="ViewBillBaseData_Car.asp?BillSn="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funBillStoreFinish(){
	error=0;
	if(!dateCheck(myForm.Sys_BillStoreFinish1.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(!dateCheck(myForm.Sys_BillStoreFinish2.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(error==0){
		UrlStr="StoreMailStationReturn.asp?day1="+myForm.Sys_BillStoreFinish1.value+"&day2="+myForm.Sys_BillStoreFinish2.value;
		newWin2(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
	}
}
function funBillStoreFinish2(){
	error=0;
	if(!dateCheck(myForm.Sys_BillStoreFinish1.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(!dateCheck(myForm.Sys_BillStoreFinish2.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(error==0){
		UrlStr="StoreMailStationReturn2.asp?day1="+myForm.Sys_BillStoreFinish1.value+"&day2="+myForm.Sys_BillStoreFinish2.value;
		newWin2(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
	}
}

function funStoreAndSendNumber(){
	error=0;
	if(!dateCheck(myForm.Sys_BillStoreFinish1.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(!dateCheck(myForm.Sys_BillStoreFinish2.value)){
		error=1;
		alert("日期輸入不正確!!");
	}
	if(error==0){
		UrlStr="StoreAndSendNumber.asp?day1="+myForm.Sys_BillStoreFinish1.value+"&day2="+myForm.Sys_BillStoreFinish2.value+"&SeqNo="+myForm.Sys_StoreAndSendNumber.value;
		newWin2(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
	}
}

function funUpdate(SN){
	UrlStr="../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funDel(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="Del";
	myForm.submit();
}
function funBillIimagePrint(StyleType){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=StyleType;
		funsubmit();
	}
}
function funBillNoPrint(StyleType){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=StyleType;
		myForm.Sys_Print.value='';
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle="+StyleType+"&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		funsubmit();
	}
}
function funBillNoPrintStyle(StyleType){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=StyleType;
		myForm.Sys_Print.value='3';
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle="+StyleType+"&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		funsubmit();
	}
}
function funFastPostReceive_HuaLien(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="FastPostReceive_style3.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funReporTrailertSend(){
	UrlStr="ReporTrailertSendList_Excel.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funMailNotBackList(){
	UrlStr="SendMailStyle.asp";
	newWin(UrlStr,"SendMailStyle",650,600,50,50,"yes","no","no","no");
}

function funMailNotBackList_TakeCar(){
	UrlStr="SendMailStyleTakeCar.asp";
	newWin(UrlStr,"SendMailStyleTakeCar",600,360,200,200,"no","no","no","no");
}

function funSealBillNo(){
	if(confirm('是否要在舉發單上顯示警員印章!!')){
		myForm.billprintuseimage.value=1;<%
			If session("Unit_ID") = "8J00" or session("Unit_ID") = "08A7" or session("Unit_ID") = "8H00" Then%>
			UrlStr="SetJobName.asp";
			myForm.action=UrlStr;
			myForm.target="Kao";
			myForm.submit();
			myForm.action="";
			myForm.target="";
		<%end if%>
	}else{
		myForm.billprintuseimage.value=0;
		myForm.hd_BillJobName.value="";
		myForm.hd_MainChName.value="";
	}
}

function funBillNoTel(){
	if(confirm('是否要在舉發單上顯示陳情電話!!')){
		myForm.Sys_BillPrintUnitTel.value='(07)7452001';
	}else{
		myForm.Sys_BillPrintUnitTel.value='';
	}
}

function funFastPostReceive(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funFastPostReceiveA4(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_1.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function label_Style_TaiNaNCity(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="label_Style_TaiNai.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLabelStyle(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="label_Style.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelStylePingtung(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="label_Style_Pingtung.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelFormat(){
	newWin("SendStyle.asp","SendKind",400,200,50,10,"yes","yes","yes","no");
}
function funLabelFormat_New(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");

		UrlStr="label_Style_Keelung_New.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelFormat_Deliver(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		<%If Session("UnitLevelID")>1 Then%>
			//if(confirm("是否要套印分局地址?")){myForm.Sys_UnitLabelKind.value='2'}
			myForm.Sys_UnitLabelKind.value='2';
		<%end if%>

		UrlStr="label_Style_Keelung_Deliver.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelFormat_Update(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");

		UrlStr="label_Style_Keelung_Update.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelFormat_act(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");

		UrlStr="BillPrintLegal_KeeLung_010911_act.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelStyleKeelung(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		if (!winopen.closed){winopen.close();}
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		UrlStr="label_Style_Keelung.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function label_Style_Keelung_NanTou(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="label_Style_Keelung_NanTou.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function label_Style_Keelung_CHCG(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="label_Style_Keelung_CHCG.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelStyleKeelung_TaiChungCity(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		UrlStr="label_Style_Keelung.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLabelStyleKeelung_miaoli(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		UrlStr="PasserUrge_miaoli_021129_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="miaoli";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function PrintPicture_HuaLien(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="BillPrintImage_HuaiLien_1081114.asp";
		myForm.action=UrlStr;
		myForm.target="BillPrintImage_HuaiLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLabelStyleLabel_miaoli(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="PasserUrge_miaoli_021129_LabelList.asp";
		myForm.action=UrlStr;
		myForm.target="miaoli";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLegalPrintMend_KaoHsiungHarBor(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="BillPrintLegal_KaoHsiungCity.asp";

		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLegalPrintMend_KaoHsiungMend(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="BillPrintLegal_KaoHsiung_Mend.asp";

		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funLabelStyleKeelung_KaoHsiungHarBor(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="label_Style_KaoHsiungHarBor.asp";
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=17&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funLabelStyleKeelung_KaoHsiung(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="label_Style_KaoHsiung.asp";
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=17&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillNoKaoHsiungPrint(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		myForm.printStyle.value=17;
		runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="BillPrintsKaoHsiung_a4.asp";
		myForm.action=UrlStr;
		myForm.target="Keelung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funFastPostReceive_tc(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style5.asp";
		myForm.action=UrlStr;
		myForm.target="tc";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funFastPostReceive_new(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style2.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funFastPostReceive_TaiChung(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style5_TaiChung.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funFastPostReceive_TaiChungCity(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style5.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funFastPostReceive_TaiChung2(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style6_TaiChung.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funFastPostReceive_TaiChungCity2(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Style6.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funFastPostReceive_Keelung(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="FastPostReceive_Keelung.asp";
		myForm.action=UrlStr;
		myForm.target="CHGH";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funPasserUrgeHuaLien_DeliverListLabel(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeHuaLien_DeliverListLabel.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funPasserUrgetaichung_DeliverListLabel(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");

		UrlStr="PasserUrgetaichung_DeliverListLabel.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funPasserUrgetaichung_Deliver_chromat(){
	if (myForm.PBillSN.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		runServerScript("BillNoPrint.asp?SQLstr=<%=strSQL2%>&printStyle=99&chk_MailNumKind=<%=chk_MailNumKind%>&Sys_BatchNumber=<%=request("Sys_BatchNumber")%>");

		UrlStr="PasserUrgetaichung_DeliverListLabel_chromat.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeDeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillSendLegalNew(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeDeliverListNew.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillSendB5(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		//UrlStr="PasserUrgeDeliverList.asp";
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeHuaLien_DeliverListV.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funBillSendB_A4(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		//UrlStr="PasserUrgeDeliverList.asp";
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeHuaLien_DeliverListV_A4.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou2";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funBillSendB5H(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeDeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		//UrlStr="PasserUrgeDeliverList.asp";
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeHuaLien_DeliverListH.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}

function funBillSend_TaiTung(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="BillBaseHuaLien_TaiTung_DeliverFList.asp";
		myForm.action=UrlStr;
		myForm.target="TaiTung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillNonTouSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeNanTou_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillNonTouNewSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="PasserUrgeNanTou_New_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="NanTou";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillTaiChungSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeTaiChung_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="TaiChung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillCHCGLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeNanTou_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeCHCG_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="CHCG";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillHuaLienSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeHuaLien_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funBillChiayiCitySendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeChiayiCity_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="ChiayiCity";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funBillTaiTungSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeTaiTung_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="TaiTung";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funBillKaoHsiungSendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//runServerScript("BillNoPrint_KaoHsiung.asp?SQLstr=<%=strSQL2%>");
		UrlStr="PasserUrgeKaoHsiung_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funBillKaoHsiungCitySendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		UrlStr="PasserUrgeKaoHsiungCity_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
}
function funBillTaiChungCitySendLegal(){
	if (myForm.Sys_strSQL.value==""){
		alert("請先輸入作業批號或舉發單號查詢欲列印的舉發單！");
	}else{
		//UrlStr="PasserUrgeHuaLien_DeliverList.asp?PBillSN="+myForm.PBillSN.value;
		/*if(confirm("是否要縣市分類?")){
			myForm.Sys_CityKind.value='1';
		}*/
		UrlStr="PasserUrgeTaiChungCity_DeliverList.asp";
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		//newWin(UrlStr,"UrgeDeliver",920,600,50,10,"yes","yes","yes","no");
	}
}
function funsubmit(){
	if(myForm.printStyle.value=='0'){
		<%If Session("UnitLevelID")>1 Then%>
			//if(confirm("是否要套印分局地址?")){myForm.Sys_UnitLabelKind.value='2'}
			myForm.Sys_UnitLabelKind.value='2';
		<%end if%>		
		UrlStr="BillPrints.asp";
	}else if(myForm.printStyle.value=='2'){
		UrlStr="BillPrints_a4.asp";
	}else if(myForm.printStyle.value=='1'){
		UrlStr="BillPrints_legalA4.asp";
	}else if(myForm.printStyle.value=='3'){
		UrlStr="BillImagePrint.asp";
	}else if(myForm.printStyle.value=='4'){
		//UrlStr="BillPrints_lattice.asp";
		UrlStr="BillPrintLegal_YiLan.asp";
	}else if(myForm.printStyle.value=='5'){
		UrlStr="BillPrints_lattice_MU.asp";
	}else if(myForm.printStyle.value=='6'){
		UrlStr="BillPrints_lattice_NanTou.asp";
	}else if(myForm.printStyle.value=='7'){
		UrlStr="BillPrintsCHCG_a4.asp";
	}else if(myForm.printStyle.value=='8'){
		UrlStr="BillPrints_lattice_HuaLien2.asp";
	}else if(myForm.printStyle.value=='9'){
		UrlStr="BillPrintsYunLin_a4.asp";
	}else if(myForm.printStyle.value=='10'){
		UrlStr="BillPrintsChiayi_a4.asp";
	}else if(myForm.printStyle.value=='11'){
		UrlStr="BillPrints_lattice_TaiChung.asp";
	}else if(myForm.printStyle.value=='12'){
		//UrlStr="BillPrints_lattice_TaiChungCity.asp";
		UrlStr="BillPrintsTaiChungCity_a4.asp";
	}else if(myForm.printStyle.value=='13'){
		UrlStr="BillPrints_lattice_YiLan.asp";
	}else if(myForm.printStyle.value=='14'){
		UrlStr="BillPrintsTaiNanCity_a4.asp";
	}else if(myForm.printStyle.value=='15'){
		//UrlStr="BillPrintsTaiTung_legalA4.asp";
		UrlStr="BillPrintsTaiTung_legalA4_img.asp";
	}else if(myForm.printStyle.value=='16'){
		UrlStr="BillPrints_PingTung.asp";
	}else if(myForm.printStyle.value=='17'){
		UrlStr="BillPrintsKaoHsiung_a4.asp";
	}else if(myForm.printStyle.value=='18'){
		UrlStr="BillPrints_TaiChungCity.asp";
	}else if(myForm.printStyle.value=='19'){
		UrlStr="BillPrints_TaiChungCity_new.asp";
	}else if(myForm.printStyle.value=='20'){
		UrlStr="BillPrintsKaoHsiung_a4_Mend.asp";
	}else if(myForm.printStyle.value=='21'){
		UrlStr="BillPrints_a4_ChiayiCity.asp";
	}else if(myForm.printStyle.value=='22'){
		UrlStr="BillPrintLegal_YiLan.asp";
	}else if(myForm.printStyle.value=='23'){
		UrlStr="BillPrintLegal_YiLan_new.asp";
	/*smith for 雲林縣使用.  */
	}else if(myForm.printStyle.value=='24'){
		UrlStr="BillPrints_TaiChung.asp";
	}else if(myForm.printStyle.value=='25'){
		UrlStr="BillPrintsTaiTung_new.asp";

	}else if(myForm.printStyle.value=='26'){
		<%If Session("UnitLevelID")>1 Then%>
			//if(confirm("是否要套印分局地址?")){myForm.Sys_UnitLabelKind.value='2'}
			myForm.Sys_UnitLabelKind.value='2';
		<%end if%>
		UrlStr="BillPrintLegal_Keelung.asp";

	}else if(myForm.printStyle.value=='27'){
		UrlStr="BillPrintsKaoHsiungHarBor_a4.asp";

	}else if(myForm.printStyle.value=='271'){
		UrlStr="BillPrintsKaoHsiungHarBor_a42.asp";

	}else if(myForm.printStyle.value=='28'){

		if(confirm("要使用43元國內郵資嗎?")){
			myForm.Sys_UnitLabelKind.value='2';
		} else {
			myForm.Sys_UnitLabelKind.value='';
		}

		UrlStr="BillPrintLegal_YiLan_new_chromat.asp";

	}else if(myForm.printStyle.value=='29'){
		UrlStr="BillPrintLegal_KaoHsiungCity_001109.asp";

	}else if(myForm.printStyle.value=='30'){
		UrlStr="BillPrintLegal_HuaLien_010213.asp";

	}else if(myForm.printStyle.value=='31'){
		UrlStr="BillPrintLegal_KaoHsiungCity_noimage_001109.asp";
	}else if(myForm.printStyle.value=='32'){
		UrlStr="BillPrintsChiayi_a4_010905.asp";

	}else if(myForm.printStyle.value=='33'){
		<%If Session("UnitLevelID")>1 Then%>
			//if(confirm("是否要套印分局地址?")){myForm.Sys_UnitLabelKind.value='2'}
			myForm.Sys_UnitLabelKind.value='2';
		<%end if%>
		UrlStr="BillPrintLegal_KeeLung_010911.asp";

	}else if(myForm.printStyle.value=='34'){
		UrlStr="BillPrintsTaiTung_new_chromat_1020321.asp";
	
	}else if(myForm.printStyle.value=='35'){
		UrlStr="BillPrintLegal_KaoHsiungCity_020902_uit.asp";

	}else if(myForm.printStyle.value=='36'){
		UrlStr="BillPrintLegal_KaoHsiungCity_noimage_020902_uit.asp";

	}else if(myForm.printStyle.value=='37'){
		UrlStr="BillPrintA4_miaoli_021129.asp";

	}else if(myForm.printStyle.value=='38'){
		UrlStr="BillPrintsCHCG_a4_1030217.asp";

	}else if(myForm.printStyle.value=='39'){
		UrlStr="BillPrints_a4_penghu030501.asp";

	}else if(myForm.printStyle.value=='40'){
		UrlStr="BillPrintsChiayi_a4_030529.asp";

	}else if(myForm.printStyle.value=='41'){
		UrlStr="BillBaseFastPaper_miaoli.asp";

	}else if(myForm.printStyle.value=='42'){
		UrlStr="BillPrints_ChiayiCity_a4_1030715.asp";

	}else if(myForm.printStyle.value=='43'){
		UrlStr="BillPrintLegal_Chiayi_1031027.asp";

	}else if(myForm.printStyle.value=='44'){
		UrlStr="BillPrintLegal_Chiayi_noimage_1031027.asp";
	
	}else if(myForm.printStyle.value=='45'){
		UrlStr="BillPrints_TaiChungCity_1040526.asp";

	}else if(myForm.printStyle.value=='46'){
		UrlStr="BillPrints_TaiChungCity_1050503.asp";

	}else if(myForm.printStyle.value=='47'){
		UrlStr="BillPrints_TaiChungCity_1050829.asp";

	}else if(myForm.printStyle.value=='48'){
		UrlStr="BillPrints_a4_kma_1060315.asp";
	
	}else if(myForm.printStyle.value=='49'){
		UrlStr="BillPrints_a4_kma_1030501.asp";
	
	}else if(myForm.printStyle.value=='50'){
		UrlStr="BillPrintsA4_TaiChungCity_1060614.asp";

	}else if(myForm.printStyle.value=='51'){
		UrlStr="BillPrintLegal_PingTung_1070412.asp";

	}else if(myForm.printStyle.value=='52'){
		UrlStr="BillPrints_TaiChungCity_1070104.asp";

		
	}else if(myForm.printStyle.value=='53'){
		UrlStr="BillPrintLegal_CHCG_1070430.asp";

	}else if(myForm.printStyle.value=='54'){
		UrlStr="BillPrintLegal_PingTung_1070412.asp";

	}else if(myForm.printStyle.value=='55'){
		UrlStr="BillPrints_TaiChungCity_1071001.asp";

	}else if(myForm.printStyle.value=='56'){
		UrlStr="BillPrintLegal_KaoHsiungCity_1071108.asp";

	}else if(myForm.printStyle.value=='57'){
		UrlStr="BillPrintLegal_KaoHsiungCity_uit_1071108.asp";

	}else if(myForm.printStyle.value=='58'){
		UrlStr="BillPrintLegal_CHCG_1070430_uit.asp";

	}else if(myForm.printStyle.value=='59'){
		UrlStr="BillPrintLegal_SPHS02_1081001.asp";
	
	}else if(myForm.printStyle.value=='60'){
		UrlStr="BillPrintsSPHS01_a4_1081001.asp";

	}else if(myForm.printStyle.value=='61'){
		UrlStr="BillPrintLegal_miaoli_1081009.asp";

	}else if(myForm.printStyle.value=='62'){
		UrlStr="BillPrintLegal_KMA_1081007.asp";

	}else if(myForm.printStyle.value=='63'){
		UrlStr="BillPrintLegal_HuaLien_Stop_1080710.asp";

	}else if(myForm.printStyle.value=='64'){
		UrlStr="BillPrintLegal_HuaLien_People_1081121.asp";

	}else if(myForm.printStyle.value=='65'){
		UrlStr="BillPrintLegal_KMA_noImage_1081007.asp";

	/*smith for 高縣停管使用 */		
	}else if(myForm.printStyle.value=='76'){
		UrlStr="BillPrintLegal_KaoHsiung_StopCar.asp";
	}else if(myForm.printStyle.value=='77'){
		UrlStr="BillPrintLegal_YunLin_new.asp";		
	/*smith for 台東使用 */
	}else if(myForm.printStyle.value=='78'){
		UrlStr="BillPrintsTaiTung_legal_new.asp";								
	/*smith for 屏東縣 郵簡 使用 BillPrints_PingTungNew.asp */		
	}else if(myForm.printStyle.value=='79'){
		UrlStr="BillPrintLegal_PenTung_new.asp";	
	}else if(myForm.printStyle.value=='80'){
		UrlStr="BillPrintLegal_NanTou.asp";
	}else if(myForm.printStyle.value=='81'){
		UrlStr="BillPrintLegal_KaoHsiungCity.asp";
	}else if(myForm.printStyle.value=='82'){
		UrlStr="BillPrintLegal_NanTou_000810.asp";
	}else if(myForm.printStyle.value=='83'){
		UrlStr="BillPrintLegal_NanTou_noimage_010904.asp";
	/*smith 1020510  for 南投新格式舉發單含相片使用 */
	/*smith for 南投新格式舉發單含相片使用 */	
	}else if(myForm.printStyle.value=='98'){
		UrlStr="BillPrintLegal_NanTou_image_011210.asp";
	}else if(myForm.printStyle.value=='101'){
		UrlStr="BillPrintLegal_NanTou_000810_1.asp";
	}else if(myForm.printStyle.value=='102'){
		UrlStr="BillPrintLegal_NanTou_noimage_010904_1.asp";
	}else if(myForm.printStyle.value=='179'){
		UrlStr="BillPrintsPTimage.asp";
	}else if(myForm.printStyle.value=='270'){
		UrlStr="BillPrintLegal_PenTung_new2013.asp";
	
	}
	/*myForm.target="mainFrame";
	myForm.submit();
	myForm.action="";
	myForm.target="";*/
	/*myForm.btnprint.disabled=false;
	if(myForm.Sys_Print.value!=''){
		myForm.hd_PrintSum.value=parseInt(myForm.hd_PrintSum.value)+parseInt(myForm.Sys_Print.value);
		if(parseInt(myForm.hd_PrintSum.value)-parseInt(myForm.Sys_Print.value)>chkcnt){
			myForm.btnprint.disabled=true;
		}else{
			myForm.btnprint.disabled=false;
		}
	}
	setTimeout('',2000);
	newWin(UrlStr,"JudeBat",920,600,50,10,"yes","yes","yes","no");*/
	myForm.action=UrlStr;
	myForm.target="JudeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	/*if(myForm.printStyle.value!='4'){
		setTimeout('funchgprint()',4000);
	}*/
}
function funUrgeList(){
	UrlStr="JudeStyle.asp";
	newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
	myForm.action="JudeStyle.asp";
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funJudesubmit(){
	winopen.close();
	if(myForm.printStyle.value=='0'){
		UrlStr="BillPrints_legal.asp";		
		newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
		myForm.action=UrlStr;
		myForm.target="UrgeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		setTimeout('funchgprint()',2000);
	}else{
		UrlStr="PasserJudeA4.asp?PBillSN="+myForm.PBillSN.value;
		newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
	}
}
function funchgprint(){
	winopen.printWindow(true,5.08,5.08,5.08,5.08);
}
function funchgExecel(){
	UrlStr="DCIExchangeQry_Execel.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"inputWin",700,550,50,10,"yes","yes","yes","no");
}
function funPrintDetail(){
	UrlStr="PictureDetail.htm";
	newWin(UrlStr,"inputWin",1000,800,50,10,"yes","yes","yes","no");
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
			alert("請先輸入作業批號查詢欲列印大宗郵件清冊的舉發單！");
	}else{
		UrlStr="MailSendList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailSendList",300,125,200,100,"no","no","no","no");
	}
}
//大宗郵件2
function funMailList2(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印交寄大宗函件的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>&MailSendType=S";
		newWin(UrlStr,"MailReportList",400,250,350,200,"no","no","no","no");
	}
}
//大宗郵件2(苗栗)
function funMailListML2(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印交寄大宗函件的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>&MailSendType=SM";
		newWin(UrlStr,"MailReportList",400,250,350,200,"no","no","no","no");
	}
}
//郵費清單
function funMailMoneyList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印郵費單的舉發單！");
	}else{
		UrlStr="MailMoneyList_Select.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"MailMoneyList",300,220,350,200,"no","no","no","no");
	}
}
//逕舉
function funReportSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印逕舉移送清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReportSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReportSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}
//逕舉_花蓮A3版
function funReportSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印逕舉移送清冊的舉發單！");
	}else{
		UrlStr="ReportSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin2",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停
function funStopSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="StopSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="StopSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停高雄市
function funStopSendList_KSC(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
		if (sys_City=='苗栗縣'){
			UrlStr="StopSendList_Excel_KSC.asp?SQLstr=<%=tmpSQL%>&Selt_MemberStation="+myForm.Selt_MemberStation.value;
		}else{
			UrlStr="StopSendList_Excel_KSC.asp?SQLstr=<%=tmpSQL%>";
		}
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停苗栗
function funStopSendList_ML(){
	//if (myForm.DB_Display.value==""){
	//		alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	//}else{

		UrlStr="StopSendList_Set_ML.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin3A",800,700,0,0,"yes","yes","yes","no");
	//}
}
//攔停A3_高雄市
function funStopSendList_A3_KSC(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
		UrlStr="StopSendList_Excel_A3_KSC.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//攔停_花蓮A3
function funStopSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印攔停移送清冊的舉發單！");
	}else{
		UrlStr="StopSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin3",800,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊
function funValidSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
	<%if sys_City="基隆市" then%>
		UrlStr="ValidSendList_Excel_GL.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ValidSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin4",800,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊_苗栗
function funValidSendList_ML(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Set_ML.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4M",700,400,0,0,"yes","yes","yes","no");
	}
}
//
function funValidSendList_ML_Stop(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Excel_ML_Stop.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4M",1000,700,0,0,"yes","yes","yes","no");
	}
}
//有效清冊_花蓮A3版
function funValidSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印有效清冊的舉發單！");
	}else{
		UrlStr="ValidSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin4",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊
function funUselessSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊_花蓮A3版
function funUselessSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//無效清冊_苗栗版
function funUselessSendList_ML(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="UselessSendList_Excel_ML.asp?SQLstr=<%=strwhere_G8ML%>";
		newWin(UrlStr,"inputWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//結案清冊
function funCaseCloseSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="CaseCloseSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"CaseCloseWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//結案清冊_花蓮A3版
function funCaseCloseSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印無效清冊的舉發單！");
	}else{
		UrlStr="CaseCloseSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"CaseCloseWin5",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(全部)
function funReturnSendList_Store_All(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Store_All.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Store_All.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Store_All.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(未結案)
function funReturnSendList_Store(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Store.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Store.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Store.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(已結案)
function funReturnSendList_Store_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Store_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin41",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(未結案)
function funReturnSendList_Gov_All(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Gov_All.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Gov_All.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="南投縣" then%>
		UrlStr="ReturnSendList_Excel_Gov_All_NT.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Gov_All.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(未結案)
function funReturnSendList_Gov(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="南投縣" then%>
		UrlStr="ReturnSendList_Excel_Gov_NT.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Gov.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(已結案)
function funReturnSendList_Gov_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="ReturnSendList_Excel_CH_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="ReturnSendList_Excel_A3_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="南投縣" then%>
		UrlStr="ReturnSendList_Excel_Gov_Close_NT.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="ReturnSendList_Excel_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin65",800,700,0,0,"yes","yes","yes","no");
	}
}
//======================================
//退件清冊_寄存(未結案)
function funReturnSendList_Store_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Store.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_寄存(已結案)A4
function funReturnSendList_Store_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Store_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin41",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(未結案)A4
function funReturnSendList_Gov_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Gov.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//退件清冊_公示(已結案)A4
function funReturnSendList_Gov_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="ReturnSendList_Excel_Gov_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin65",800,700,0,0,"yes","yes","yes","no");
	}
}
//================================================
//收受
function funGetSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="GetSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="GetSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//收受_花蓮A3版
function funGetSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印退件清冊的舉發單！");
	}else{
		UrlStr="GetSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin6",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊
function funStoreSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin7",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊_花蓮A3版
function funStoreSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
		UrlStr="funStoreSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin7",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(未結案)
function funStoreSendList_UnClose(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="funStoreSendList_Excel_A3_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin71",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(未結案)_A4版
function funStoreSendList_UnClose_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
		UrlStr="funStoreSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin71",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(已結案)
function funStoreSendList_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="funStoreSendList_Excel_A3_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funStoreSendList_Excel_CH_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funStoreSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin72",800,700,0,0,"yes","yes","yes","no");
	}
}
//寄存送達清冊(已結案)_A4版
function funStoreSendList_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印寄存送達清冊的舉發單！");
	}else{
		UrlStr="funStoreSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin72",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊
function funGovSendList(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(不分頁)
function funGovSendList_NoPage(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel_NoPage.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊_花蓮
function funGovSendList_HL(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(已結案)
function funGovSendList_Close(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="funGovSendList_Excel_A3_Close.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH_Close.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin81",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(已結案)_A4版
function funGovSendList_Close_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel_Close.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin81",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(未結案)
function funGovSendList_UnClose(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
	<%if sys_City="花蓮縣" or sys_City="嘉義縣" or sys_City="嘉義市" then%>
		UrlStr="funGovSendList_Excel_A3_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%elseif sys_City="彰化縣" then '彰化不要數量統計表而且每頁筆數較少%>
		UrlStr="funGovSendList_Excel_CH_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%else%>
		UrlStr="funGovSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
	<%end if%>
		newWin(UrlStr,"inputWin82",800,700,0,0,"yes","yes","yes","no");
	}
}
//公示送達清冊(未結案)_A4版
function funGovSendList_UnClose_A4(){
	if (myForm.DB_Display.value==""){
			alert("請先輸入作業批號查詢欲列印公示送達清冊的舉發單！");
	}else{
		UrlStr="funGovSendList_Excel_UnClose.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin82",800,700,0,0,"yes","yes","yes","no");
	}
}
//車籍查詢
function funchgCarDataList(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印車籍清冊的舉發單！");
	}else{
		UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","no","yes","no");
	}
}
//車籍查詢_花蓮A3版
function funchgCarDataList_HL(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印車籍清冊的舉發單！");
	}else{
		UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhereToPrintCarData%>";
		newWin(UrlStr,"DciCarListWin",790,575,50,10,"yes","no","yes","no");
	}
}
//公告清冊
function funOpenGovList(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印公告清冊的舉發單！");
	}else{
		UrlStr="funOpenGovList_A3.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
//不郵寄清冊
function funNotMailList(){
	if (myForm.DB_Display.value==""){
		alert("請先輸入作業批號查詢欲列印不郵寄清冊的舉發單！");
	}else{
		UrlStr="NotMailList_Excel.asp?SQLstr=<%=tmpSQL%>";
		newWin(UrlStr,"inputWin8",800,700,0,0,"yes","yes","yes","no");
	}
}
function funBillMailInfoMark(){
	myForm.action="BillMailInfoMark.asp";
	myForm.target="inputWin2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
</script>
<%conn.close%>