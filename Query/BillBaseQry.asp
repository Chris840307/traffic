<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單管理</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%
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

BillBaseName="BillBaseView"
'組成查詢SQL字串
if request("DB_Selt")="Selt" Then
	strQry="[查詢]"
	if sys_City="花蓮縣" then 
		strwhere=" and a.ImagePathName is null "
	else
		strwhere=""
	end if
	TCRecorddate_check=0
		if trim(request("IllegalDateCheck"))="1" then
			if request("IllegalDate")<>"" and request("IllegalDate1")<>""Then
				strQry=strQry&"IllegalDate="&Trim(request("IllegalDate"))&",IllegalDate1="&Trim(request("IllegalDate1"))
				ArgueDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
				ArgueDate2=gOutDT(request("IllegalDate1"))&" 23:59:59"
				strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
			end if
		end if
		if trim(request("RecordDateCheck"))="1" then
			if request("RecordDate")<>"" and request("RecordDate1")<>""Then
				If strQry="[查詢]" then
					strQry=strQry&"RecordDate="&Trim(request("RecordDate"))&",RecordDate1="&Trim(request("RecordDate1"))
				Else
					strQry=strQry&",RecordDate="&Trim(request("RecordDate"))&",RecordDate1="&Trim(request("RecordDate1"))
				End if
				if sys_City="台中市" then 
					Recdate1Temp=DateAdd("m",-3,now)
					if DateDiff("d",gOutDT(request("RecordDate")),Recdate1Temp)>0 then
						TCRecorddate_check=1
					end if 
				end if 
				RecordDate1=gOutDT(request("RecordDate"))&" 0:0:0"
				RecordDate2=gOutDT(request("RecordDate1"))&" 23:59:59"
				if strwhere<>"" then
					strwhere=strwhere&" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				else
					strwhere=" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				end if
			end if
		end if
		if trim(request("BillFillDateCheck"))="1" then
			if request("BillFillDate")<>"" and request("BillFillDate1")<>""Then
				If strQry="[查詢]" then
					strQry=strQry&"BillFillDate="&Trim(request("BillFillDate"))&",BillFillDate1="&Trim(request("BillFillDate1"))
				Else
					strQry=strQry&",BillFillDate="&Trim(request("BillFillDate"))&",BillFillDate1="&Trim(request("BillFillDate1"))
				End if
				BillFillDate1=gOutDT(request("BillFillDate"))&" 0:0:0"
				BillFillDate2=gOutDT(request("BillFillDate1"))&" 23:59:59"
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillFillDate between TO_DATE('"&BillFillDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&BillFillDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				else
					strwhere=" and a.BillFillDate between TO_DATE('"&BillFillDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&BillFillDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				end if
			end if
		end if
		if trim(request("RecordDate_h"))<>"" or trim(request("RecordDate1_h"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			else
				strwhere=" and to_char(a.RecordDate,'hh') between "&trim(request("RecordDate_h"))&" and "&trim(request("RecordDate1_h"))
			end if
		end if
		'有輸入建檔人，且建檔人=操作者時，忽略舉發單位
		if trim(request("Sys_RecordMemberID"))="" or trim(session("User_ID"))<>trim(request("Sys_RecordMemberID")) then
			if request("Sys_BillUnitID")<>"" Then
				If strQry="[查詢]" then
					strQry=strQry&"BillUnitID="&Trim(request("Sys_BillUnitID"))
				Else
					strQry=strQry&",BillUnitID="&Trim(request("Sys_BillUnitID"))
				End If
				if sys_City="台中市" And (Trim(request("Sys_BillUnitID"))="0460" Or Trim(request("Sys_BillUnitID"))="0406" Or Trim(request("Sys_BillUnitID"))="0410" Or Trim(request("Sys_BillUnitID"))="0420" Or Trim(request("Sys_BillUnitID"))="0430" Or Trim(request("Sys_BillUnitID"))="0440" Or Trim(request("Sys_BillUnitID"))="0450" Or Trim(request("Sys_BillUnitID"))="0480" Or Trim(request("Sys_BillUnitID"))="4A00" Or Trim(request("Sys_BillUnitID"))="4B00" Or Trim(request("Sys_BillUnitID"))="4C00" Or Trim(request("Sys_BillUnitID"))="4D00" Or Trim(request("Sys_BillUnitID"))="4E00" Or Trim(request("Sys_BillUnitID"))="4F00" Or Trim(request("Sys_BillUnitID"))="4G00" Or Trim(request("Sys_BillUnitID"))="4H00" Or Trim(request("Sys_BillUnitID"))="4I00") then 
					strwhere=strwhere&" and a.BillUnitID in (select UnitID from UnitInfo where UnitTypeID='"&request("Sys_BillUnitID")&"')"
				else	
					strwhere=strwhere&" and a.BillUnitID in ('"&request("Sys_BillUnitID")&"')"
				End If 
			end if
		end if 
		if trim(request("Sys_BillMem"))<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"BillMem="&Trim(request("Sys_BillMem"))
			Else
				strQry=strQry&",BillMem="&Trim(request("Sys_BillMem"))
			End if
			strwhere=strwhere&" and (a.BillMemID1 in (Select memberID from MemberData where LoginID='"&trim(request("Sys_BillMem"))&"')"
			strwhere=strwhere&" or a.BillMemID2 in (Select memberID from MemberData where LoginID='"&trim(request("Sys_BillMem"))&"')"
			strwhere=strwhere&" or a.BillMemID3 in (Select memberID from MemberData where LoginID='"&trim(request("Sys_BillMem"))&"')"
			strwhere=strwhere&" or a.BillMemID4 in (Select memberID from MemberData where LoginID='"&trim(request("Sys_BillMem"))&"'))"

		end if
		'smith for taichicity
	if sys_City = "台中市" then 	
	
	elseif sys_City = "苗栗縣" and Session("UnitLevelID")=1 then 	

		if request("Sys_RecordUnit")<>"" and request("Sys_RecordMemberID")="" Then
			if instr(request("Sys_RecordUnit"),",")=0 then

				If strQry="[查詢]" then
					strQry=strQry&"Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
				Else
					strQry=strQry&",Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
				End If 

				strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
			end if
		end If 
	elseif sys_City = "苗栗縣" then





	elseif sys_City = "台南市" then 	
		if request("Sys_RecordUnit")<>"" and request("Sys_RecordMemberID")="" Then
			if instr(request("Sys_RecordUnit"),",")=0 then
				If strQry="[查詢]" then
					strQry=strQry&"Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
				Else
					strQry=strQry&",Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
				End if
				strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
			end if
		end if
	else
		if request("Sys_RecordUnit")<>"" and request("Sys_RecordMemberID")="" Then
			If strQry="[查詢]" then
				strQry=strQry&"Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
			Else
				strQry=strQry&",Sys_RecordUnit="&Trim(request("Sys_RecordUnit"))
			End if
			strwhere=strwhere&" and a.RecordMemberID in (select MemberID from MemberData where UnitID in ('"&trim(request("Sys_RecordUnit"))&"'))"
		end if
	end if			
		if request("Sys_RecordMemberID")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"RecordMemberID="&Trim(request("Sys_RecordMemberID"))
			Else
				strQry=strQry&",RecordMemberID="&Trim(request("Sys_RecordMemberID"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and a.RecordMemberID ="&request("Sys_RecordMemberID")
			else
				strwhere=" and a.RecordMemberID="&request("Sys_RecordMemberID")
			end if
		end if
		if request("Sys_BillTypeID")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"BillTypeID="&Trim(request("Sys_BillTypeID"))
			Else
				strQry=strQry&",BillTypeID="&Trim(request("Sys_BillTypeID"))
			End if
			if trim(request("Sys_BillTypeID"))="1" then
				strSys_BillTypeID=" and a.BillBaseTypeID='0' and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
			elseif trim(request("Sys_BillTypeID"))="2" then
				strSys_BillTypeID=" and a.BillBaseTypeID='0' and a.BillTypeID='"&request("Sys_BillTypeID")&"'"
			elseif trim(request("Sys_BillTypeID"))="3" then
				strSys_BillTypeID=" and a.BillBaseTypeID='1'"
			elseif trim(request("Sys_BillTypeID"))="9" then
				strSys_BillTypeID=" and a.BillBaseTypeID='0'"
			end if

			if strwhere<>"" then
				strwhere=strwhere&strSys_BillTypeID
			else
				strwhere=strSys_BillTypeID
			end if
		end if
		if request("Sys_BillNo")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"BillNo="&Trim(request("Sys_BillNo"))
			Else
				strQry=strQry&",BillNo="&Trim(request("Sys_BillNo"))
			End If
			If InStr(Trim(request("Sys_BillNo")),",")>0 Then
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillNo in ('"&Replace(request("Sys_BillNo"),",","','")&"')"
				else
					strwhere=" and a.BillNo in ('"&Replace(request("Sys_BillNo"),",","','")&"')"
				end if
			Else
				if strwhere<>"" then
					strwhere=strwhere&" and a.BillNo='"&request("Sys_BillNo")&"'"
				else
					strwhere=" and a.BillNo='"&request("Sys_BillNo")&"'"
				end if
			End If 			
		end if
		if request("Sys_CarNo")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"CarNo="&Trim(request("Sys_CarNo"))
			Else
				strQry=strQry&",CarNo="&Trim(request("Sys_CarNo"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and a.CarNo='"&UCase(request("Sys_CarNo"))&"'"
			else
				strwhere=" and a.CarNo='"&UCase(request("Sys_CarNo"))&"'"
			end if
		end if

		if request("Sys_DriverID")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"DriverID="&Trim(request("Sys_DriverID"))
			Else
				strQry=strQry&",DriverID="&Trim(request("Sys_DriverID"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and a.DriverID='"&request("Sys_DriverID")&"'"
			else
				strwhere=" and a.DriverID='"&request("Sys_DriverID")&"'"
			end if
		end if
		if request("Sys_LawID")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"Rule1="&Trim(request("Sys_LawID"))
			Else
				strQry=strQry&",Rule1="&Trim(request("Sys_LawID"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and (a.Rule1 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule2 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule3 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule4 like '"&request("Sys_LawID")&"%')"
			else
				strwhere=" and (a.Rule1 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule2 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule3 like '"&request("Sys_LawID")&"%'" &_
					" or a.Rule4 like '"&request("Sys_LawID")&"%')"
			end if
		 end If
		 If sys_City = "高雄市" then 
			if request("Sys_StreetID1")<>"" Then
				If strQry="[查詢]" then
					strQry=strQry&"IllegalAddressID="&Trim(request("Sys_StreetID1"))
				Else
					strQry=strQry&",IllegalAddressID="&Trim(request("Sys_StreetID1"))
				End if
				if strwhere<>"" then
					strwhere=strwhere&" and a.IllegalAddressID='"&request("Sys_StreetID1")&"'"
				else
					strwhere=" and a.IllegalAddressID='"&request("Sys_StreetID1")&"'"
				end if
			end if
		End If 
		If sys_City = "台南市" then 
			if request("Sys_Foreigner")="1" Then
				if strwhere<>"" then
					strwhere=strwhere&" and ( not (substr(driverid,1,1) between 'A' and 'Z' and substr(driverid,2,1) between '1' and '2') and driverid is not null)"
				else
					strwhere=" and ( not (substr(driverid,1,1) between 'A' and 'Z' and substr(driverid,2,1) between '1' and '2') and driverid is not null)"
				end If
			End If 
		End If 
		
		if request("Sys_StreetID")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"IllegalAddress="&Trim(request("Sys_StreetID"))
			Else
				strQry=strQry&",IllegalAddress="&Trim(request("Sys_StreetID"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and a.IllegalAddress like '%"&request("Sys_StreetID")&"%'"
			else
				strwhere=" and a.IllegalAddress like '%"&request("Sys_StreetID")&"%'"
			end if
		end if
		if request("DCIstatus")<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"BillStatus="&Trim(request("DCIstatus"))
			Else
				strQry=strQry&",BillStatus="&Trim(request("DCIstatus"))
			End if
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='"&request("DCIstatus")&"'"
			else
				strwhere=" and a.BillStatus='"&request("DCIstatus")&"'"
			end if
		end if
		if trim(request("Sys_TrafficAccidentNo"))<>"" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.TrafficAccidentNo='"&request("Sys_TrafficAccidentNo")&"'"
			else
				strwhere=" and a.TrafficAccidentNo='"&request("Sys_TrafficAccidentNo")&"'"
			end if
		end if
'		if trim(request("CaseClose"))<>"" then
'			if trim(request("CaseClose"))="1" then
'				strwhere=strwhere&" and a.BillStatus='9'"
'			elseif trim(request("CaseClose"))="2" then
'				strwhere=strwhere&" and a.BillStatus<>'9'"
'			end if
'		end if
		
		if trim(request("RecordStateID"))="close" then
			if strwhere<>"" then
				strwhere=strwhere&" and a.BillStatus='9'"
			else
				strwhere=" and a.BillStatus='9'"
			end if
			If strQry="[查詢]" then
				strQry=strQry&"RecordStateID=結案"
			Else
				strQry=strQry&",RecordStateID=結案"
			End if
		elseif trim(request("RecordStateID"))="NotCaseIn" then
				strwhere=strwhere&" and a.BillStatus in ('0','1') and a.RecordStateID=0"
				If strQry="[查詢]" then
					strQry=strQry&"RecordStateID=未入案"
				Else
					strQry=strQry&",RecordStateID=未入案"
				End if
		elseif trim(request("RecordStateID"))="1" then
				strwhere=strwhere&" and not exists(select 'Y' from dcilog where exchangetypeid='W' and billsn=a.sn) and exists(select 'Y' from dcilog where exchangetypeid='A' and billsn=a.sn)"
				If strQry="[查詢]" then
					strQry=strQry&"RecordStateID=車籍查詢"
				Else
					strQry=strQry&",RecordStateID=車籍查詢"
				End if
		elseif trim(request("RecordStateID"))="all" then
			'99/5/3基隆攔停陳與陳朝招決定攔停刪除未入案不出現
			if sys_City="基隆市" then 
				strwhere=strwhere&" and (a.RecordStateID=0 or (a.BillBaseTypeID='0' and a.BillTypeID='2' and a.BillNo is not null and a.RecordStateID=-1) or (a.BillBaseTypeID='0' and a.BillTypeID='1' and a.BillNo is not null and a.RecordStateID=-1 and a.Sn in (select a.BillSN from Dcilog a,DciReturnStatus b where a.ExchangeTypeID='W' and a.ExchangeTypeID=b.DciActionID and a.DciReturnStatusid=b.DciReturn and b.DciReturnStatus=1)) or (a.BillBaseTypeID='1' and a.BillNo is not null and a.RecordStateID=-1))"
			else
				strwhere=strwhere&" and (a.RecordStateID=0 or (a.BillNo is not null and a.RecordStateID=-1))"
			end if
			If strQry="[查詢]" then
				strQry=strQry&"RecordStateID=全部"
			Else
				strQry=strQry&",RecordStateID=全部"
			End if
		elseif trim(request("RecordStateID"))="0" then
			strwhere=strwhere&" and a.RecordStateID=0"
			If strQry="[查詢]" then
				strQry=strQry&"RecordStateID=有效"
			Else
				strQry=strQry&",RecordStateID=有效"
			End if
		elseif trim(request("RecordStateID"))="-1" then
			if sys_City="基隆市" then 
				strwhere=strwhere&" and ((a.BillBaseTypeID='0' and a.BillTypeID='2' and a.BillNo is not null and a.RecordStateID=-1) or (a.BillBaseTypeID='0' and a.BillTypeID='1' and a.BillNo is not null and a.RecordStateID=-1 and a.Sn in (select a.BillSN from Dcilog a,DciReturnStatus b where a.ExchangeTypeID='W' and a.ExchangeTypeID=b.DciActionID and a.DciReturnStatusid=b.DciReturn and b.DciReturnStatus=1)) or (a.BillBaseTypeID='1' and a.BillNo is not null and a.RecordStateID=-1))"
			else
				strwhere=strwhere&" and a.BillNo is not null and a.RecordStateID=-1"
			end if
			If strQry="[查詢]" then
				strQry=strQry&"RecordStateID=已刪除"
			Else
				strQry=strQry&",RecordStateID=已刪除"
			End if
		end if 
		if trim(request("DoubleChk"))<>"" then
				strwhere=strwhere&" and a.DoubleCheckStatus="&trim(request("DoubleChk"))
		end If
		
		if trim(request("MailNumber"))<>"" Then
			BillBaseName="BillBase"
			If strQry="[查詢]" then
				strQry=strQry&"MailNumber="&Trim(request("MailNumber"))
			Else
				strQry=strQry&",MailNumber="&Trim(request("MailNumber"))
			End if
				strwhere=strwhere&" and (c.MailNumber like '%"&trim(request("MailNumber"))&"%' or c.StoreAndSendMailNumber like '%"&trim(request("MailNumber"))&"%')"
		end if
		if trim(request("PeopleMailNumber"))<>"" Then
			BillBaseName="PasserBase"
			If strQry="[查詢]" then
				strQry=strQry&"PeopleMailNumber="&Trim(request("PeopleMailNumber"))
			Else
				strQry=strQry&",PeopleMailNumber="&Trim(request("PeopleMailNumber"))
			End if
				strwhere=strwhere&" and exists(select SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN=a.SN and SendMailStation like '%"&trim(request("PeopleMailNumber"))&"%')"
		end if
		if trim(request("BatchNumber"))<>"" Then
			BillBaseName="BillBase"
			If strQry="[查詢]" then
				strQry=strQry&"BatchNumber="&Trim(request("BatchNumber"))
			Else
				strQry=strQry&",BatchNumber="&Trim(request("BatchNumber"))
			End if
				strwhere=strwhere&" and a.SN in (select BillSN from DciLog where BatchNumber='"&trim(request("BatchNumber"))&"')"
		end if	
		If Trim(request("ReportNo"))<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"ReportNo="&Trim(request("ReportNo"))
			Else
				strQry=strQry&",ReportNo="&Trim(request("ReportNo"))
			End if
				strwhere=strwhere&" and exists (select 'Y' from (select billsn from BillReportNo where ReportNo='"&trim(request("ReportNo"))&"' union all select billsn from PasserReportNo where ReportNo='"&trim(request("ReportNo"))&"') tmpa where billsn=a.sn)"
		End If 
		
		If sys_City = "澎湖縣" then
			If Trim(request("chkNoImagePH"))="1" Then
				strwhere=strwhere&" and not exists (select BillNo from BillAttatchImage where BillNo=a.BillNo and TypeID=1 and Recordstateid=0) "
			End If 
		End If 
		'If sys_City = "金門縣" then
			If Trim(request("chkTrafficAccidentType"))="1" Then
				If Trim(request("TrafficAccidentType"))="" Then
					strwhere=strwhere&" and exists(select Sn from billbase where Sn=a.Sn and TrafficAccidentType in ('1','2','3')) "
				Else
					strwhere=strwhere&" and exists(select Sn from billbase where Sn=a.Sn and TrafficAccidentType='"&Trim(request("TrafficAccidentType"))&"') "
				End If 
			End If 
		'End If 

		If sys_City = "彰化縣" then
			If Trim(request("Sys_Name"))<>"" Then
				strwhere=strwhere&" and exists(select CarNo from billbaseDciReturn where BillNo=a.BillNo and CarNo=a.CarNo and ExchangeTypeID='W' and (Owner='"&trim(request("Sys_Name"))&"' or Driver='"&trim(request("Sys_Name"))&"')) "
			End If 
		End If 

		if trim(request("sys_OrderType"))="1" or trim(request("sys_OrderType"))="" then
			strOrderPlus="a.RecordDate"
		elseif trim(request("sys_OrderType"))="2" then
			strOrderPlus="a.IllegalDate"
		elseif trim(request("sys_OrderType"))="3" then
			strOrderPlus="a.BillNo"
		elseif trim(request("sys_OrderType"))="4" then
			strOrderPlus="c.UserMarkDate"
		elseif trim(request("sys_OrderType"))="5" then
			strOrderPlus="a.BillUnitID"
		end if

		if trim(request("sys_OrderType2"))="1" or trim(request("sys_OrderType2"))="" then
			strOrderPlus2=""
		elseif trim(request("sys_OrderType2"))="2" then
			strOrderPlus2=" Desc"
		end if

		',a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID
		strColPlus=""
		If sys_City = "台中市" then 
			strColPlus=",(select IllegalZip from billbase where sn=a.sn) as IllegalZip"
		End if
		'smith 催繳加一個欄位ImageFileNameB
  	SelField = " a.Note,a.ImageFileNameB,a.DriverID, a.SN,a.IllegalDate,a.CarSimpleID,a.BIllMemID1,a.BIllMemID2,a.BIllMemID3,a.BIllMemID4,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillBaseTypeID,a.DoubleCheckStatus,a.BillFillDate,c.MailNumber"&strColPlus&" from "&BillBaseName&" a,MemberData b,BillMailHistory c"
				
		
		If sys_City="高雄市" Then
			strSQL="select " & SelField & " where a.RecordMemberID=b.MemberID(+) and a.sn=c.BillSN(+)"&strwhere&" and rownum<=5000 order by "&strOrderPlus & strOrderPlus2 
		else
			strSQL="select " & SelField & " where a.RecordMemberID=b.MemberID(+) and a.sn=c.BillSN(+)"&strwhere&" order by "&strOrderPlus & strOrderPlus2 
		End If 
		'response.write strSQL
end if
'花連停管催繳單號
if request("DB_Selt")="Selt_Stop" then		
		BillBaseName="BillBase"
		'smith mark 讓只做到車籍查詢的也可以查
		'and a.ImageFileNameB is not null  
		strQry="[查詢]"
		strwhere=" and a.ImagePathName is not null  and a.BillNo is null and a.recordstateid=0"
		
		if trim(request("StopBillNo"))<>"" And trim(request("StopBillNo"))<>"0000000000" Then
			Sys_StopBillNo=Right("00000000000000000"&Trim(request("StopBillNo")),16)
			If strQry="[查詢]" then
				strQry=strQry&"StopBillNo="&Sys_StopBillNo
			Else
				strQry=strQry&",StopBillNo="&Sys_StopBillNo
			End if
			strwhere=strwhere&" and a.ImageFileNameB='"&Sys_StopBillNo&"'"
		end if

		if trim(request("StopCarNo"))<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"CarNo="&Trim(request("StopCarNo"))
			Else
				strQry=strQry&",CarNo="&Trim(request("StopCarNo"))
			End if
			strwhere=strwhere&" and a.CarNo='"&trim(request("StopCarNo"))&"'"
		end if

		if trim(request("sys_OrderType"))="1" or trim(request("sys_OrderType"))="" then
			strOrderPlus="a.RecordDate"
		elseif trim(request("sys_OrderType"))="2" then
			strOrderPlus="a.IllegalDate"
		elseif trim(request("sys_OrderType"))="3" then
			strOrderPlus="a.BillNo"
		elseif trim(request("sys_OrderType"))="4" then
			strOrderPlus="c.UserMarkDate"
		end if


		if trim(request("sys_OrderType2"))="1" or trim(request("sys_OrderType2"))="" then
			strOrderPlus2=""
		elseif trim(request("sys_OrderType2"))="2" then
			strOrderPlus2=" Desc"
		end if


		if trim(request("BatchNumber"))<>"" Then
			If strQry="[查詢]" then
				strQry=strQry&"BatchNumber="&Trim(request("BatchNumber"))
			Else
				strQry=strQry&",BatchNumber="&Trim(request("BatchNumber"))
			End if
			strwhere=strwhere&" and a.SN in (select BillSN from DciLog where BatchNumber='"&trim(request("BatchNumber"))&"')"
		end if	
		
		'smith 催繳加一個欄位ImageFileNameB
  		SelField = " a.recordstateid,a.Note,a.ImagePathName,a.ImageFileNameB,a.DriverID, a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillBaseTypeID,a.DoubleCheckStatus,a.BillFillDate,c.MailNumber from "&BillBaseName&" a,MemberData b,BillMailHistory c"
				
		strSQL="select " & SelField & " where a.RecordMemberID=b.MemberID(+) and a.sn=c.BillSN(+)"&strwhere&" order by "&strOrderPlus & strOrderPlus2
		
end if

'刪除未入案舉發單
if trim(request("kinds"))="Del_NoDci" then
	DelMemID=trim(Session("User_ID"))
	theBillSN=trim(request("Del_SN"))

	'抓單號
	theBillNO=""
	theCarNO=""
	strbillno="select BillNo,CarNo from BillBase where SN="&theBillSN
	set rsBillno=conn.execute(strbillno)
	if not rsBillno.eof then
		theBillNO=trim(rsBillno("BillNo"))
		theCarNO=trim(rsBillno("CarNo"))
	end if
	rsBillno.close
	set rsBillno=nothing

		NoteTmp=""
			
		'該筆紀錄的打驗資料表的 BILLSTATUS 更新為 6
		strUpdDelTemp="Update BillBaseTmp set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where BillNo='"&theBillNO&"'"
		conn.execute strUpdDelTemp

		'更新該筆紀錄的 BILLSTATUS 更新為 6
		strUpdDel="Update BillBase set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&theBillSN
		conn.execute strUpdDel

		DeleteReason="無"
		ConnExecute "舉發單刪除 單號:"&theBillNO&" 車號:"&theCarNO&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
end if

'做完車籍查詢及入案等動作後再查詢告發單，讓列表取得的資料為最新
if request("DB_Selt")="Selt" or request("DB_Selt")="Selt_Stop" then
'response.write strQry
'response.write "<br>"
'response.end
		'response.write strSQL
		'response.end
		

		set rsfound=conn.execute(strSQL)
		
		strCnt="select count(*) as cnt from "&BillBaseName&" a,MemberData b,BillMailHistory c where a.RecordMemberID=b.MemberID(+) and a.sn=c.BillSN(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close

		'ConnExecute DBsum,399
		strForChkMen=""
		If sys_City="高雄市" then
			strForChkMen="，代查詢人:"&Trim(request("ForChkMen"))
		end if 
		If sys_City="屏東縣" or sys_City="基隆市" or sys_City="台東縣" then
			if Trim(request("ForChkMen"))="" then
				ForChkMen_Temp=Trim(session("Ch_Name"))
			else
				ForChkMen_Temp=Trim(request("ForChkMen") &"")
			end if 
			Crt_Log ForChkMen_Temp,Trim(request("QryReason")),DBsum&"筆"&strForChkMen&"，"&strQry ,355,"",""
		else
			ConnExecute DBsum&"筆"&strForChkMen&"，查詢事由:"&Trim(request("QryReason"))&"="&strQry ,355
		end if 
		tmpSQL=strwhere
		Session.Contents.Remove("BillSQL")
		Session("BillSQL")=strSQL
		Session.Contents.Remove("PrintCarDataSQL")
		Session("PrintCarDataSQL")=strwhere
		Session.Contents.Remove("PrintCarDataSQLCheckItem")
		Session("PrintCarDataSQLCheckItem")=strQry

		
end if


%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">
<%  'smith 20091017高雄市的刪除都稱為註銷

	if sys_City="高雄市" Or sys_City=ApconfigureCityName then 
		sDeleteSymbol="註銷"
	else 
		sDeleteSymbol="刪除"
	end if

%>
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">舉發單管理</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
					<%if sys_City="台中市" or sys_City="高雄市" then%>
						<font style="color: #FF0000;font-size: 28px;line-height:32px;"><strong>※僅可查詢三個月內建檔案件，如有輸入單號可不勾選查詢日期<br>
						如需查詢年度資料請使用統計報表</strong></font>
						<br>
					<%end if%>
						<input type="checkbox" name="IllegalDateCheck" value="1" <%
						if trim(request("IllegalDateCheck"))="1" Or Trim(request("OpenPageFlag") & "")="" then
							response.write "checked"
						end if
						%>>
						違規日期
						<input name="IllegalDate" type="text" value="<%
						If Trim(request("OpenPageFlag") & "")="" Then
							if sys_City="澎湖縣" Or sys_City="南投縣" then
								response.write Year(DateAdd("yyyy",-1,now))-1911&Right("00" & Month(DateAdd("yyyy",-1,now)),2)&Right("00" & Day(DateAdd("yyyy",-1,now)),2)
							elseif sys_City="台中市" Or sys_City="高雄市" then
								response.write Year(DateAdd("m",-2,now))-1911&Right("00" & Month(DateAdd("m",-2,now)),2)&Right("00" & Day(DateAdd("m",-2,now)),2)
							else
								response.write Year(DateAdd("m",-6,now))-1911&Right("00" & Month(DateAdd("m",-6,now)),2)&Right("00" & Day(DateAdd("m",-6,now)),2)
							End If 
						Else
							response.write request("IllegalDate")
						End If 
						%>" size="6" maxlength="7" class="btn1" onKeyup="this.value=this.value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate');">
						~
						<input name="IllegalDate1" type="text" value="<%
						If Trim(request("OpenPageFlag") & "")="" Then
							response.write Year(now)-1911&Right("00" & Month(now),2)&Right("00" & Day(now),2)
						Else
							response.write request("IllegalDate1")
						End If 
						%>" size="6" maxlength="7" class="btn1" onKeyup="this.value=this.value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate1');">
						<!-- <img src="space.gif" width="8" height="10">
						單位 -->
						<%'=SelectUnitOption("Sys_BillUnitID","Sys_BillMem")
						%>
						<!-- <img src="space.gif" width="8" height="10">
						舉發員警 -->
						<%'=SelectMemberOption("Sys_BillUnitID","Sys_BillMem")
						%>
						<img src="space.gif" width="5" height="2">
						舉發單類別
						<select name="Sys_BillTypeID" class="btn1">
							<option Value="">全部</option>
							<option value="1" <%if trim(request("Sys_BillTypeID"))="1" then response.write "selected"%>>攔停</option>
							<option value="2" <%if trim(request("Sys_BillTypeID"))="2" then response.write "selected"%>>逕舉</option>
							<option value="3" <%if trim(request("Sys_BillTypeID"))="3" then response.write "selected"%>>慢車行人道路障礙</option>
						<%If Session("Credit_ID")="A000000000" then%>
							<option value="9" <%if trim(request("Sys_BillTypeID"))="9" then response.write "selected"%>>攔停+逕舉</option>
						<%End if%>
						</select>
						
						<img src="space.gif" width="3" height="2">
						狀態
						<select name="RecordStateID">
					<%if sys_City="苗栗縣" then%>
							<option value="0" <%if trim(request("RecordStateID"))="0" then response.write "selected"%>>有效</option>
							<option value="all" <%if trim(request("RecordStateID"))="all" then response.write "selected"%>>全部</option>
							
					<%else%>
							<option value="all" <%if trim(request("RecordStateID"))="all" then response.write "selected"%>>全部</option>
							<option value="0" <%if trim(request("RecordStateID"))="0" then response.write "selected"%>>有效</option>
					<%End if%>
							<option value="-1" <%if trim(request("RecordStateID"))="-1" then response.write "selected"%>>已<%=sDeleteSymbol%></option>
							<option value="close" <%if trim(request("RecordStateID"))="close" then response.write "selected"%>>結案/繳費</option>
							<option value="NotCaseIn" <%if trim(request("RecordStateID"))="NotCaseIn" then response.write "selected"%>>未入案</option>
					<%if sys_City="台南市" then%>
							<option value="1" <%if trim(request("RecordStateID"))="1" then response.write "selected"%>>車籍查詢</option>
					<%End if%>
						</select>
						<!-- <img src="space.gif" width="5" height="2">
						結案
						<select name="CaseClose">
							<option value="">請選擇</option>
							<option value="1" <%if trim(request("CaseClose"))="1" then response.write "selected"%>>是</option>
							<option value="2" <%if trim(request("CaseClose"))="2" then response.write "selected"%>>否</option>
						</select> -->
						<!-- <br>
						
						
						時段
						<input name="RecordDate_h" type="text" value="<%=request("RecordDate_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時 ~ 
						<input name="RecordDate1_h" type="text" value="<%=request("RecordDate1_h")%>" size="1" maxlength="2" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">時
						<img src="space.gif" width="8" height="10"> -->					

						<img src="space.gif" width="2" height="10">
						
						<strong>舉發單號</strong>
						<input name="Sys_BillNo" type="text" value="<%=request("Sys_BillNo")%>" size="19" class="btn1" onkeyup="value=value.toUpperCase()">
					<%if sys_City = "澎湖縣" then%>
						&nbsp; &nbsp; &nbsp; 
						<input type="checkbox" name="chkNoImagePH" value="1" <%
						If Trim(request("chkNoImagePH"))="1" Then
							response.write "checked"
						End if
						%>>未上傳違規相片掃描檔案件
					<%End if%>
					<%'if sys_City="金門縣" then%>
						&nbsp; &nbsp; 
						<input type="checkbox" value="1" name="chkTrafficAccidentType" <%If Trim(request("chkTrafficAccidentType"))="1" Then response.write "checked"%>>交通事故種類
						
						
						<select name="TrafficAccidentType">
							<option value="" >全部</option>
							<option value="1" <%If Trim(request("TrafficAccidentType"))="1" Then response.write "selected"%>>A1</option>
							<option value="2" <%If Trim(request("TrafficAccidentType"))="2" Then response.write "selected"%>>A2</option>
							<option value="3" <%If Trim(request("TrafficAccidentType"))="3" Then response.write "selected"%>>A3</option>
						</select>
					<%'End if%>
						<br>
						<input type="checkbox" name="RecordDateCheck" value="1" <%
						if trim(request("RecordDateCheck"))="1" then
							response.write "checked"
						elseif (sys_City="台中市" or sys_City="高雄市") and trim(request("Sys_BillNo"))="" and trim(request("Sys_CarNo"))="" then
							response.write "checked"
						end if
						%>>
						建檔日期
						<input name="RecordDate" type="text" value="<%
						If Trim(request("OpenPageFlag") & "")="" Then
							if sys_City="台中市" or sys_City="高雄市" then
								response.write Year(DateAdd("m",-1,now))-1911&Right("00" & Month(DateAdd("m",-1,now)),2)&Right("00" & Day(DateAdd("m",-1,now)),2)
							End If 
						Else
							response.write request("RecordDate")
						End If 
						%>" size="6" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate');">
						~
						<input name="RecordDate1" type="text" value="<%
						If Trim(request("OpenPageFlag") & "")="" Then
							if sys_City="台中市" or sys_City="高雄市" then
								response.write Year(now)-1911&Right("00" & Month(now),2)&Right("00" & Day(now),2)
							End If 
						Else
							response.write request("RecordDate1")
						End If 
						%>" size="6" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
						<img src="space.gif" width="8" height="10">
	
						法條代碼
						<input name="Sys_LawID" type="text" value="<%=request("Sys_LawID")%>" size="8" class="btn1">
												
						<img src="space.gif" width="8" height="10">
						<strong>車號</strong>
						<input name="Sys_CarNo" type="text" value="<%=request("Sys_CarNo")%>" size="7" maxlength="8" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="8" height="10">
						證號
						<input name="Sys_DriverID" type="text" value="<%=request("Sys_DriverID")%>" size="10" maxlength="20" class="btn1" onkeyup="value=value.toUpperCase()">
						&nbsp; &nbsp; &nbsp; 
						<%If sys_City = "彰化縣" then%>
						姓名
						<input name="Sys_Name" type="text" value="<%=request("Sys_Name")%>" size="10" maxlength="20" class="btn1" >
						<img src="space.gif" width="8" height="10">
						<%End if%>
						<%If sys_City = "台南市" then%>
						外國人<input type="checkbox" name="Sys_Foreigner" value="1" <%
						If Trim(request("Sys_Foreigner"))="1" Then
							response.write "checked"
						End If 
						%>>
						<%End if%>
						<br>
						<input type="checkbox" name="BillFillDateCheck" value="1" <%
						if trim(request("BillFillDateCheck"))="1" then
							response.write "checked"
						end if
						%>>
						填單日期
						<input name="BillFillDate" type="text" value="<%=request("BillFillDate")%>" size="6" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('BillFillDate');">
						~
						<input name="BillFillDate1" type="text" value="<%=request("BillFillDate1")%>" size="6" maxlength="7" class="btn1"  onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('BillFillDate1');">
						<img src="space.gif" width="8" height="10">
						舉發單位
						<%'SelectUnitOption("Sys_BillUnitID","")%>
						<select name="Sys_BillUnitID" class="btn1"><%
							strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUnit=conn.execute(strSQL)
							If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
							rsUnit.close
							strUnitID=""
							if trim(Session("UnitLevelID"))="1" Then
								if sys_City="台中市" then
									strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder"
								Else
									strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
								End If	
								strtmp=strtmp+"<option value="""">所有單位</option>"
							elseif trim(Session("UnitLevelID"))="2" Then
								if sys_City="台中市" then
									strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitTypeID=(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitOrder"
								Else
									strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitTypeID=(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"
								End If 

								set rs1=conn.execute(strSQL)
								while Not rs1.eof
									if trim(strUnitID)<>"" then strUnitID=trim(strUnitID)&","
									if trim(strUnitID)="" then
										strUnitID=strUnitID&trim(rs1("UnitID"))
									else
										strUnitID=strUnitID&"'"&trim(rs1("UnitID"))
									end if
									rs1.movenext
									if Not rs1.eof then strUnitID=strUnitID&"'"
								wend
								rs1.close
								if sys_City="台南市" then
									strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
								else
									strtmp=strtmp+"<option value="""">所有單位</option>"
								end if 				
								
							elseif trim(Session("UnitLevelID"))="3" Then
								if sys_City="高雄市" then
									strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitTypeID=(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"') order by UnitOrder"
								Else
									strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitTypeID,UnitName"
								End If 
								
							end if
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
								if trim(rs1("UnitID"))=trim(request("Sys_BillUnitID")) then
									strtmp=strtmp+" selected"
								end if
								strtmp=strtmp+">"&rs1("UnitName")&"</option>"
								rs1.movenext
							wend
							rs1.close
							strtmp=strtmp+"</select>"
							response.write strtmp%>
						<img src="space.gif" width="8" height="10">
						舉發員警代碼
						<input type="text" name="Sys_BillMem" class="btn1" value="<%=trim(request("Sys_BillMem"))%>" size="5">

						<br>
						<img src="space.gif" width="20" height="22">
						建檔單位
						<%=UnSelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						建檔人
						<%=UnSelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="5" height="10">
					<%If sys_City = "高雄市" then%>
						違規地點代碼
						<input name="Sys_StreetID1" type="text" value="<%=request("Sys_StreetID1")%>" size="5" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="5" height="10">
					<%End if%>
						違規地點
						<input name="Sys_StreetID" type="text" value="<%=request("Sys_StreetID")%>" size="21" class="btn1">
						<img src="space.gif" width="5" height="10">
						<!-- 事故案件編號
						<input name="Sys_TrafficAccidentNo" type="text" value="<%=request("Sys_TrafficAccidentNo")%>" size="10" maxlength="15" class="btn1">
						<img src="space.gif" width="8" height="10"> -->
						
				<%	DoubleChkFlag=0
					strDoubleChk="select Value from ApConfigure where ID=38"
					set rsDoubleChk=conn.execute(strDoubleChk)
					if rsDoubleChk.eof then
						DoubleChkFlag=trim(rsDoubleChk("Value"))
					end if
					rsDoubleChk.close
					set rsDoubleChk=nothing
					if DoubleChkFlag=1 then
				%>
						<img src="space.gif" width="8" height="10">
						一打一驗
						<select name="DoubleChk">
							<option value="">請選擇</option>
							<option value="0" <%if trim(request("DoubleChk"))="0" then response.write "selected"%>>未通過</option>
							<option value="1" <%if trim(request("DoubleChk"))="1" then response.write "selected"%>>通過</option>
							<option value="-1" <%if trim(request("DoubleChk"))="-1" then response.write "selected"%>>需更正</option>
						</select>
				<%end if%>
						<br>
						<img src="space.gif" width="19" height="22">
						大宗掛號碼
						<input type="text" class="btn1" style="width: 200px;" value="<%=trim(request("MailNumber"))%>" name="MailNumber" >
						
						
						作業批號
						<input type="text" class="btn1" size="10" value="<%=trim(request("BatchNumber"))%>" name="BatchNumber" onkeyup="value=value.toUpperCase();">
				<%If sys_City="台中市" then%>
						<img src="space.gif" width="3" height="22">
						告示單號
						<input type="text" class="btn1" size="8" value="<%=trim(request("ReportNo"))%>" name="ReportNo" >
				<%elseIf sys_City="高雄市" or sys_City="基隆市" then%>
						<img src="space.gif" width="3" height="22">
						標示單號
						<input type="text" class="btn1" size="10" value="<%=trim(request("ReportNo"))%>" name="ReportNo" >
				<%End If %>
						<img src="space.gif" width="8" height="10">
						<!-- 慢車、微電車大宗掛號碼
						<input type="text" class="btn1" style="width: 200px;" value="<%=trim(request("PeopleMailNumber"))%>" name="PeopleMailNumber" >

						<br /> -->
						
						查詢事由
						<select name="QryReason" class="btn1">
				<%'If sys_City<>"台中市" then%>
							<option value="" >請選擇</option>
				<%'End If %>
				<%If sys_City<>"屏東縣" then%>
							<option value="資料檢核" <%If Trim(request("QryReason"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
				<%End If %>
				<%If sys_City<>"嘉義縣" then%>
							<option value="執行業務" <%If Trim(request("QryReason"))="執行業務" Then response.write "selected" End if%>>執行業務</option>
				<%End If %>
				<%If sys_City="屏東縣" then%>
							<option value="民眾申訴(來電)" <%If Trim(request("QryReason"))="民眾申訴(來電)" Then response.write "selected" End if%>>民眾申訴(來電)</option>
							<option value="民眾申訴(臨櫃)" <%If Trim(request("QryReason"))="民眾申訴(臨櫃)" Then response.write "selected" End if%>>民眾申訴(臨櫃)</option>
							<option value="民眾申訴(公文)" <%If Trim(request("QryReason"))="民眾申訴(公文)" Then response.write "selected" End if%>>民眾申訴(公文)</option>
				<%else%>
							<option value="民眾申訴" <%If Trim(request("QryReason"))="民眾申訴" Then response.write "selected" End if%>>民眾申訴</option>
				<%End If %>			
				<%If sys_City="屏東縣" then%>
							<option value="資料檢核" <%If Trim(request("QryReason"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
				<%End If %>
				<%If sys_City="嘉義縣" then%>
							<option value="定期查核" <%If Trim(request("QryReason"))="定期查核" Then response.write "selected" End if%>>定期查核</option>
				<%else%>
							<option value="事故處理" <%If Trim(request("QryReason"))="事故處理" Then response.write "selected" End if%>>事故處理</option>
				<%End If %>	
				<%If sys_City<>"台中市" and sys_City<>"屏東縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"雲林縣" and sys_City<>"嘉義市" then%>
							<option value="偵查刑案" <%If Trim(request("QryReason"))="偵查刑案" Then response.write "selected" End if%>>偵查刑案</option>
				<%End If %>
				
						</select>
				<%If sys_City="高雄市" or sys_City="屏東縣" or sys_City="台東縣" then%>
						<img src="space.gif" width="8" height="10">
						代查詢人
						<input type="text" class="btn1" size="8" value="<%=trim(request("ForChkMen"))%>" name="ForChkMen" >
				<%End If %>
						<img src="space.gif" width="8" height="22">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();" <%
						if CheckPermission(234,1)=false then
							response.write "disabled"
						end if
						%>>
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseQry.asp'">
				<%if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
						<input type="button" name="btnDelAll" value="多筆刪除"  onclick='window.open("../Query/BillBase_Del_More.asp","WebPage2_Del_All","left=200,top=100,location=0,width=460,height=380,resizable=yes,scrollbars=yes")'
						<%
'							IF trim(request("BatchNumber"))="" Then
'								response.write "disabled"
'							else
'								if checkUpLoadUser(trim(request("BatchNumber")))=false then
'									response.write "disabled"
'								end if
'							end if
						%>>
				<%end if%>
				<%If Trim(session("UnitLevelID"))<>"1" and Trim(session("UnitLevelID"))<>"2" And sys_City="台南市" then%>

				<%else%>
						<input type="button" name="btnDelAll" value="整批<%=sDeleteSymbol%>"  onclick='window.open("../Query/BillBase_Del_All.asp?BatchNumber=<%=trim(request("BatchNumber"))%>","WebPage2_Del_All","left=300,top=200,location=0,width=360,height=210,resizable=yes,scrollbars=yes")'
						<%
						if request("DB_Selt")<>"Selt" and request("DB_Selt")<>"Selt_Stop" then
							response.write "disabled"
						else
							IF trim(request("BatchNumber"))="" Then
								response.write "disabled"
							ElseIf InStr(trim(request("BatchNumber")),"W")=0 and InStr(trim(request("BatchNumber")),"A")=0 Then
								response.write "disabled"
							Else
								If sys_City="苗栗縣" then	
									If Trim(session("Credit_ID"))="A01" Or Trim(session("Credit_ID"))="A000000000" Then

									Else
										if checkUpLoadUser(trim(request("BatchNumber")))=false then
											response.write "disabled"
										end if
									End if 
								elseif checkUpLoadUser(trim(request("BatchNumber")))=false then
									response.write "disabled"
								end if
							end if
						end if
						%>> 
				<%End if%>
						<input type="button" name="btnDelAll" value="整批撤銷送達"  onclick='window.open("../Query/BillBase_Cancel_All.asp?BatchNumber=<%=trim(request("BatchNumber"))%>","WebPage2_Del_All","left=300,top=200,location=0,width=360,height=210,resizable=yes,scrollbars=yes")'
						<%
						if request("DB_Selt")<>"Selt" and request("DB_Selt")<>"Selt_Stop" then
							response.write "disabled"
						else
							IF trim(request("BatchNumber"))="" Then
								response.write "disabled"
							else
								if checkUpLoadUser(trim(request("BatchNumber")))=false then
									response.write "disabled"
								end if
							end if
						end if
						%> style="font-size: 10pt; width: 90px; height:26px;"> 
						
				<%if trim(Session("UpdateBillUser"))="1" then%>
						<input type="button" name="btnDelAll" value="整批修改"  onclick='window.open("../BillKeyIn/BillBase_Update_All.asp?BatchNumber=<%=trim(request("BatchNumber"))%>","WebPage2_Del_All","left=0,top=0,location=0,width=800,height=600,resizable=yes,scrollbars=yes")'
						<%
						if request("DB_Selt")<>"Selt" and request("DB_Selt")<>"Selt_Stop" then
							response.write "disabled"
						end if
						%>>
				<%end if%>
				<%if sys_City="花蓮縣" Or sys_City="台東縣" Or sys_City="屏東縣" then%>
						<br>
						催繳單號
						<input type="text" class="btn1" size="16" maxlength="16" value="<%
					if sys_City="花蓮縣" then
						If trim(request("StopBillNo"))<>"" then
							response.write trim(request("StopBillNo"))
						Else
							response.write "0000000000"
						End If 
					Else
						response.write trim(request("StopBillNo"))
					End If	
						%>" name="StopBillNo" onkeyup="value=value.toUpperCase();">
						催繳車號
						<input name="StopCarNo" type="text" value="<%=request("StopCarNo")%>" size="8" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">
						<input type="button" name="btStopBill" value="催繳單查詢" onclick="Selt_Stop()">
				<%end if%>
						
						
						<input type="hidden" name="OpenPageFlag" value="1">
						<br />
						<span style="color: #FF0000;font-size: 18px;"><strong>※案件建檔人員才可做刪除案件。</strong></span>
				<%'If sys_City="彰化縣" or sys_City="台南市" then%>
						<br />
						<span style="color: #FF0000;font-size: 18px;"><strong>※強制入案案件，如案件處理人員沒有人工移交給監理站建檔入案，監理站會沒有這筆舉發單。</strong></span>
				<%'end if%>
				<%If sys_City="台南市" then%>
						<br />
						<span style="color: #FF0000;font-size: 18px;line-height:20px;"><strong>※因資安審查規定，查詢時必須輸入舉發單單號。</strong></span>
				<%end if%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF" class="style3">
			舉發單紀錄列表
			<img src="space.gif" width="56" height="8">
			每頁 
			<select name="sys_MoveCnt" onchange="repage();">
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
			</select>
			筆 
			&nbsp; &nbsp; &nbsp; &nbsp; 
			<font color="#F90000"><strong>(共 <%=DBsum%> 筆 含 建檔未入案件 數)</strong></font>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			依&nbsp;<select name="sys_OrderType" onchange="repage();">
				<option value="1" <%if trim(request("sys_OrderType"))="1" then response.write " Selected"%>>建檔日期</option>
				<option value="2" <%if trim(request("sys_OrderType"))="2" then response.write " Selected"%>>違規日期</option>
				<option value="3" <%if trim(request("sys_OrderType"))="3" then response.write " Selected"%>>舉發單號</option>
				<option value="4" <%if trim(request("sys_OrderType"))="4" then response.write " Selected"%>>註記日期</option>
				<option value="5" <%if trim(request("sys_OrderType"))="5" then response.write " Selected"%>>舉發單位</option>
			</select>
			<select name="sys_OrderType2" onchange="repage();">
				<option value="1" <%if trim(request("sys_OrderType2"))="1" then response.write " Selected"%>>由小至大</option>
				<option value="2" <%if trim(request("sys_OrderType2"))="2" then response.write " Selected"%>>由大至小</option>
				
			</select>
			排列

		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th width="5%" >類別</th>
					<th width="8%" nowrap>狀態</th>
					<!--smith 催繳加一個欄位ImageFileNameB-->
					<% if sys_City="花蓮縣" then %>
						<th width="6%" nowrap>單號</th>
					<%else%>
						<th width="6%" nowrap>舉發單號</th>
					<% end if %>
					
					<th width="8%" nowrap>車號</th>
					
					<th width="7%">違規日</th>
				        <!-- <th width="6%">入案日</th> -->
					<th width="8%">舉發員警</th>
					
					
					<th width="5%">車種</th>
					
					<th width="7%">駕駛人</th>
					<th width="20%">違規地點</th>
					<th width="8%">法條</th>
				<%if DoubleChkFlag=1 then%>
					<th width="6%">一打<br>一驗</th>
				<%end if%>
				<%If sys_City="台東縣" Then%>
					<th width="6%">郵寄<br>序號</th>
				<%End If %>
					<th width="16%">操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
				if request("DB_Selt")="Selt" or request("DB_Selt")="Selt_Stop" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						chname="":chRule="":ForFeit=""
						if rsfound("BillMem1")<>"" then	chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then	chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then	chname=chname&"/"&rsfound("BillMem3")
						if rsfound("BillMem4")<>"" then	chname=chname&"/"&rsfound("BillMem4")
						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						'if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")

						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"
						
						'--------------------舉發單類別-----------------------
						response.write "<td>"
						if trim(rsfound("BillBaseTypeID"))="0" then
							strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
							set rsBTypeVal=conn.execute(strBTypeVal)
							if not rsBTypeVal.eof then
								response.write rsBTypeVal("Content")
							end if
							rsBTypeVal.close
							set rsBTypeVal=nothing
						else
							if trim(rsfound("BillTypeID"))="1" then
								response.write "慢車 – 攔停"
							elseif trim(rsfound("BillTypeID"))="2" then
								response.write "慢車 – 逕舉"
							elseif trim(rsfound("BillTypeID"))="3" then
								response.write "道路"
							end if
						end if
						response.write "</td>"
						'---------------------舉發單狀態------------------------------------	
						BillStatusTmp=""
						strBillStatus=""
						if rsfound("BillStatus")="0" then
							BillStatusTmp="<span class=""style5"">建檔</span>"
						elseif rsfound("BillStatus")="1" then
							BillStatusTmp="<span class=""style5"">查車</span>"
							strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='A' order by exchangedate desc"
						elseif rsfound("BillStatus")="2" then
							' smith 催繳 的入案狀態顯示為催繳
							if sys_City="花蓮縣" Or sys_City="台東縣" Or sys_City="屏東縣" then 
								if trim(rsfound("ImageFileNameB")) <> "" then						
									BillStatusTmp="<span class=""style5"">催繳</span>"
								else
									BillStatusTmp="<span class=""style5"">入案</span>"									
								end if								
							else
								BillStatusTmp="<span class=""style5"">入案</span>"								
							end if
							if trim(rsfound("BillBaseTypeID"))="0" then
								strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='W' order by exchangedate desc"
							else
								if trim(rsfound("CarNo"))<>"" and not isnull(rsfound("CarNo")) then	'慢車電動車
									strBillStatus="select * from PASSERDCILOG where BillSn="&rsfound("SN")&" and ExchangeTypeID='W' order by exchangedate desc"
								end if 
							end if 
						elseif rsfound("BillStatus")="3" then
							BillStatusTmp="<span class=""style5"">單退</span>"
							strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='N' and ReturnMarkType='3' order by exchangedate desc"
						elseif rsfound("BillStatus")="4" then
							BillStatusTmp="<span class=""style5"">寄存</span>"
							strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='N' and ReturnMarkType='4' order by exchangedate desc"
						elseif rsfound("BillStatus")="5" then
							BillStatusTmp="<span class=""style5"">公示</span>"
							strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='N' and ReturnMarkType='5' order by exchangedate desc"
						elseif rsfound("BillStatus")="6" then
							BillStatusTmp=sDeleteSymbol
							if trim(rsfound("BillBaseTypeID"))="0" then
								strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='E' order by exchangedate desc"
							else
								if trim(rsfound("CarNo"))<>"" and not isnull(rsfound("CarNo")) then	'慢車電動車
									strBillStatus="select * from PASSERDCILOG where BillSn="&rsfound("SN")&" and ExchangeTypeID='E' order by exchangedate desc"
								end if 
							end if 
							
						elseif rsfound("BillStatus")="7" then
							BillStatusTmp="<span class=""style5"">收受</span>"
							strBillStatus="select * from DciLog where BillSn="&rsfound("SN")&" and ExchangeTypeID='N' and ReturnMarkType='7' order by exchangedate desc"
						elseif rsfound("BillStatus")="9" then
							BillStatusTmp="<span class=""style5"">結案</span>"
							if trim(rsfound("BillBaseTypeID"))="0" then	'攔停逕舉沒結案檔

							else
								if trim(rsfound("CarNo"))<>"" and not isnull(rsfound("CarNo")) then	'慢車電動車
									strBillStatus="select * from PASSERDCILOG where BillSn="&rsfound("SN")&" and ExchangeTypeID='B' order by exchangedate desc"
								end if 
							end if 
						end if
						DelStatusID_tmp=0
						if strBillStatus<>"" then
							set rsBStatus=conn.execute(strBillStatus)
							if not rsBStatus.eof then
								if trim(rsBStatus("DcireturnStatusID"))="" or isnull(rsBStatus("DcireturnStatusID")) then
									BillStatusTmp=BillStatusTmp&".未處理"
								else
									strDci="select DciReturnStatus from DciReturnStatus " &_
										" where DciActionID='"&trim(rsBStatus("ExchangeTypeID"))&"' " &_
										" and DciReturn='"&trim(rsBStatus("DcireturnStatusID"))&"'"
									set rsDci=conn.execute(strDci)
									if not rsDci.eof then
										if rsDci("DciReturnStatus")="1" then
											DelStatusID_tmp=1
											BillStatusTmp=BillStatusTmp&"<span class=""style5"">.正常</span>"
										else
											DelStatusID_tmp=0
											BillStatusTmp=BillStatusTmp&"<span class=""style5"">.異常</span>"
										end if
									else
										DelStatusID_tmp=1
									end if
									rsDci.close
									set rsDci=nothing
								end if
							end if
							rsBStatus.close
							set rsBStatus=nothing

						end if
						response.write "<td>"&BillStatusTmp
						strDSupd="select count(billsn) as cnt from DCISTATUSUPDATE where Billsn="&Trim(rsfound("Sn"))
						Set rsDSupd=conn.execute(strDSupd)
						If Not rsDSupd.eof Then
							If CDbl(rsDSupd("cnt"))>0 Then
								response.write "<br><font style='font-size:10pt;color: #FF0000; '>( 強制入案案件 )</font>"
							End If 
						End If
						rsDSupd.close
						Set rsDSupd=nothing
						response.write "</td>"
						'-------------------------------------------------------------------
						'smith 催繳加一個欄位ImageFileNameB
						if sys_City="花蓮縣" Or sys_City="台東縣" Or sys_City="屏東縣" then 
							response.write "<td width='6%'>"&rsfound("BillNo")&rsfound("ImageFileNameB")&"</td>"
						else
							response.write "<td width='6%'>"&rsfound("BillNo")&"</td>"
						end if
						response.write "<td width='6%' nowarp>"&rsfound("CarNo")&"</td>"
						response.write "<td width='5%'><span class=""style5"">"&gInitDT(trim(rsfound("IllegalDate")))&"<br>"&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&minute(rsfound("IllegalDate")),2)&"</span></td>"
						
					'到BillBaseDciReturn抓入案日期------------------------------------------
					'目前計算出來做後續判斷. 但是不顯示
					'response.write "<td>"
					if trim(rsfound("BillBaseTypeID"))="0" then
						CaseInDate=""
						CaseINCnt=0
						strCType="select a.DciCaseInDate,a.Status from BillBaseDCIReturn a " &_
						" where ((a.BillNo='"&trim(rsfound("BillNo"))&"' and a.CarNo='"&trim(rsfound("CarNo"))&"')" &_
						" or (a.BillNo is null and a.CarNo='"&trim(rsfound("CarNo"))&"')) and ExchangeTypeID='W'" &_
						" order by DciCaseInDate Desc"
						set rsCType=conn.execute(strCType)
						if not rsCType.eof then
							'response.write trim(rsCType("DciCaseInDate"))
							if trim(rsCType("Status"))="Y" or trim(rsCType("Status"))="S" or trim(rsCType("Status"))="n" then
								if len(trim(rsCType("DciCaseInDate")))=8 then
									CaseInDate=gOutDT((left(trim(rsCType("DciCaseInDate")),4)-1911) & right(trim(rsCType("DciCaseInDate")),4))
								else
									CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
								end if 
								CaseINCnt=1
							else
								CaseINCnt=2
							end if
						else
							strCType2="select ExchangeDate,DciReturnStatusID from DciLog where BillSn='"&trim(rsfound("Sn"))&"' " &_
								" and ExchangeTypeID='W' " &_
								" order by ExchangeDate Desc"
							set rsCType2=conn.execute(strCType2)
							if not rsCType2.eof then
								if trim(rsCType2("DciReturnStatusID"))="Y" or trim(rsCType2("DciReturnStatusID"))="S" or trim(rsCType2("DciReturnStatusID"))="n" or isnull(rsCType2("DciReturnStatusID")) then
									CaseInDate=trim(rsCType2("ExchangeDate"))
									CaseINCnt=1
								else
									CaseINCnt=2
								end if
							end if
							rsCType2.close
							set rsCType2=nothing
						end if
						rsCType.close
						set rsCType=nothing
						'response.write CaseInDate
						'計算入案幾天
						CountCaseIN=0
						if CaseInDate<>"" then
							CountCaseIN=dateDiff("d",CaseInDate,now)
						end if
					else
						CaseINCnt=0
						CountCaseIN=0
					end if
					'response.write "</td>"
					'----------------------------------------------
						response.write "<td><span class=""style5"">"&chname&"</span></td>"
						
						
						
						response.write "<td>"
							if trim(rsfound("BillBaseTypeID"))="0" then
								if trim(rsfound("CarSimpleID"))="1" then
									response.write "<span class=""style5"">汽車</span>"
								elseif trim(rsfound("CarSimpleID"))="2" then
									response.write "<span class=""style5"">拖車</span>"
								elseif trim(rsfound("CarSimpleID"))="3" then
									response.write "<span class=""style5"">重機</span>"
								elseif trim(rsfound("CarSimpleID"))="4" then
									response.write "<span class=""style5"">輕機</span>"
								elseif trim(rsfound("CarSimpleID"))="5" then
									response.write "<span class=""style5"">動力機械</span>"
								elseif trim(rsfound("CarSimpleID"))="6" then
									response.write "<span class=""style5"">臨時車牌</span>"
								elseif trim(rsfound("CarSimpleID"))="7" then
									response.write "<span class=""style5"">試車牌</span>"
								end if
							end if
						response.write "</td>"

						response.write "<td>"&rsfound("Driver")&"</td>"
						response.write "<td align=""left""><span class=""style5"">"
						If sys_City = "台中市" then 
							response.write rsfound("IllegalZip")&" "
						End if
						response.write rsfound("IllegalAddress")&"</span></td>"
						response.write "<td>"&chRule&"</td>"
					if DoubleChkFlag=1 then
						response.write "<td>"
							if trim(rsfound("DoubleCheckStatus"))="0" then
								response.write "未通過"
							elseif trim(rsfound("DoubleCheckStatus"))="1" then
								response.write "&nbsp;"
							elseif trim(rsfound("DoubleCheckStatus"))="-1" then
								response.write "需更正"
							end if
						response.write "</td>"
					end If
					If sys_City="台東縣" Then
						if trim(rsfound("BillBaseTypeID"))="0" then
							response.write "<td>"&trim(rsfound("MailNumber"))&"</td>"
						else
							response.write "<td>"
							strPeopleMN="select min(SendMailStation) SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsfound("SN"))
							set rsPeopleMN=conn.execute(strPeopleMN)
							if not rsPeopleMN.eof then
								response.write trim(rsPeopleMN("SendMailStation"))
							end if 
							rsPeopleMN.close
							set rsPeopleMN=nothing 
							response.write "</td>"
						end if 
					End If 

						response.write "<td align='left'>"
						
					'----------------------------------------------------------------------------------------------						
					if trim(rsfound("BillBaseTypeID"))="0" then
						'催繳單攔停 逕舉的詳細
						if sys_City="花蓮縣" Or sys_City="台東縣" Or sys_City="屏東縣" then 
							if trim(rsfound("ImageFileNameB")) <> "" then	%>						
									<input type="button" name="b1" value="詳細" onclick='window.open("<%
								If sys_City="台東縣" Then
									response.write "StopBillBaseData_Detail_TaiDung.asp"
								Else
									response.write "StopBillBaseData_Detail.asp"
								End If 
									%>?BillSN=<%=trim(rsfound("SN"))%>&BillType=0&QryReason=<%=trim(request("QryReason"))%>&ForChkMen=<%=Trim(request("ForChkMen"))%>","WebPage2","left=0,top=0,location=0,width=980,height=555,resizable=yes,scrollbars=yes,menubar=yes,status=yes")' style="font-size: 10pt; width: 40px; height:26px;">
									<%If sys_City="屏東縣" then%>

									<%else%>
									<input type="button" name="b1" value="補印催繳單" onclick='window.open("StopBillPrints_HuaLien_Mend.asp?SQLstr= and a.sn=<%=trim(rsfound("SN"))%>&BillType=0","WebPage2","left=0,top=0,location=0,width=980,height=555,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 80px; height:26px;">
									<%end if%>	
							<%else%>
									<input type="button" name="b1" value="詳細" onclick='window.open("BillBaseData_Detail.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=0&QryReason=<%=trim(request("QryReason"))%>&ForChkMen=<%=Trim(request("ForChkMen"))%>","WebPage2","left=0,top=0,location=0,width=980,height=755,resizable=yes,scrollbars=yes,menubar=yes,status=yes")' style="font-size: 10pt; width: 40px; height:26px;"> 
							<%end if%>						
<%
					  else
					  	'攔停 逕舉的詳細%>			
							<input type="button" name="b1" value="詳細" onclick='window.open("BillBaseData_Detail.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=0&QryReason=<%=trim(request("QryReason"))%>","WebPage2","left=0,top=0,location=0,width=980,height=755,resizable=yes,scrollbars=yes,menubar=yes,status=yes")' style="font-size: 10pt; width: 40px; height:26px;"> 
<%					
						end if
					else
%>	
						<!--'慢車行人道路障礙的詳細-->
						<input type="button" name="b1" value="詳細" onclick='window.open("ViewBillBaseData_People.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=1","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;"> 
						
<%				end if
				'----------------------------------------------------------------------------------------
				CanUpdateBillData=0
				If sys_City="彰化縣" Then
					If Trim(Session("Unit_ID"))="JS00" Or Trim(Session("Unit_ID"))="JM00" Then
						strUchk="Select UnitTypeID from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
						Set rsUchk=conn.execute(strUchk)
						If Not rsUchk.eof Then
							If Trim(rsUchk("UnitTypeID"))=Trim(Session("Unit_ID")) Then
								CanUpdateBillData=1
							End If 
							'response.write Trim(rsUchk("UnitTypeID"))
						End If 
						rsUchk.close
						Set rsUchk=Nothing 
					End If 
				End If 
				if trim(rsfound("RecordStateID"))="0" then
					'非刪除之舉發單才能作操作
					'使用者=建檔人才可以做修改刪除
					'雲林花蓮系統管理員也能改，苗栗只有A01能改
					'保二總隊四大隊二中隊=南科 ，只有二中隊能改
					if ((trim(rsfound("RecordMemberID"))=trim(session("User_ID")) and sys_City<>"雲林縣" and sys_City<>"花蓮縣" and sys_City<>"保二總隊四大隊二中隊") or ((trim(rsfound("RecordMemberID"))=trim(session("User_ID")) or Session("Group_ID")="200") and (sys_City="雲林縣" or sys_City="花蓮縣" or sys_City="連江縣"))) Or (sys_City="苗栗縣" And Trim(session("Credit_ID"))="A01") or (Trim(Session("Unit_ID"))="R01" and sys_City="保二總隊四大隊二中隊") Or (sys_City="台中市" And Trim(session("Credit_ID"))="Z016") Or CanUpdateBillData=1 then
						if trim(rsfound("BillBaseTypeID"))="0" then
							if trim(rsfound("BillTypeID"))="2" then	
							'逕舉
	%>	
							<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes,status=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
	<%							
							else	
							'攔停
							'高雄市攔停已結案
								if (sys_City="高雄市" Or sys_City=ApconfigureCityName) and trim(rsfound("BillStatus"))="9" and (trim(rsfound("BillUnitID"))="0861" or trim(rsfound("BillUnitID"))="0862" or trim(rsfound("BillUnitID"))="0863" or trim(rsfound("BillUnitID"))="0864" or trim(rsfound("BillUnitID"))="0871") then%>
							<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_TakeCar_Update.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
	<%							else%>
							<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_Car_Update.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
	<%							end if
							end if
						else	'passerbase 
							
							chkPB_Edit=0
							PBaseUrl="../BillKeyIn/BillKeyIn_People.asp?BillSN="&trim(rsfound("SN"))
							
							strSQL="select count(1) cnt from PasserBase where Driver is not null and sn="&trim(rsfound("SN"))

							set rspb=conn.execute(strSQL)
							chkPB_Edit=cdbl(rspb("cnt"))
							rspb.close
							
'							If chkPB_Edit = 0 Then
'
'								strSQL="select count(1) cnt from PASSERDCILOG where exchangetypeid='W' and billsn="&trim(rsfound("SN"))
'
'								set rspb=conn.execute(strSQL)
'								chkPB_Edit=cdbl(rspb("cnt"))
'								rspb.close
'							End if 
							
							If chkPB_Edit = 0 Then PBaseUrl=PBaseUrl&"&hid_BillTypeID=2"

	%>	
							<input type="button" name="b1" value="修改" onclick='window.open("<%=PBaseUrl%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes,status=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
	<%							
						end if

					elseif trim(rsfound("BillBaseTypeID"))="0" then
						'BillBase登入者<>建檔人，但(session("ManagerPower"))="1" 
						If sys_City="苗栗縣" And Trim(Session("Unit_ID"))<>"03BA" Then
						
						else
	%>					<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_Dci_Update.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
	<%					End If 
					end If
					If sys_City="苗栗縣" And Trim(Session("Unit_ID"))="03BA" Then
				%>
						<input type="button" name="b1" value="修改地址" onclick='window.open("../BillKeyIn/BillKeyIn_Dci_Update.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes,status=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 60px; height:26px;">
				<%
					End If 
					'============================================
					' 花蓮縣用 imagefilenameb 以及 rule1 是不是56 + note 裡面是不是有催繳的檔案.txt副檔名 來判斷出現刪除按鈕													
				if (checkIsAllowDel(sys_City,trim(rsfound("BillTypeID")))=true) or (trim(rsfound("imagefilenameb"))<>"")  or ( (Instr(rsfound("Rule1"),"56")>0) and (Instr(rsfound("Note"),"txt")>0) and (sys_City="花蓮縣") ) Or (sys_City="苗栗縣" And Trim(session("Credit_ID"))="A01") Or (Session("Group_ID")="200" and sys_City="連江縣") Or (sys_City="台中市" And Trim(session("Credit_ID"))="Z016") Or CanUpdateBillData=1 or Trim(session("Credit_ID"))="A000000000" or (sys_City="保二總隊四大隊二中隊" and (Trim(session("Credit_ID"))="19870107")) then
					if sys_City="保二總隊四大隊二中隊" and (Trim(session("Credit_ID"))<>"19870107" ) then
						'南科 統一由1987刪除
					else
						if trim(rsfound("BillBaseTypeID"))="0" then	'BillBase
							'入案日期超過七天不能刪除
							if CountCaseIN<8 then
								'1:查詢 ,2:新增 ,3:修改 ,4:刪除
								if CheckPermission(234,4)=true then
									'未入案直接刪
									if CaseINCnt=51315 Then '避免使用者先開舉發單管理,再入案,然後刪除,所以不能直接刪
										if sys_City="台中市" then
		%>
								<input type="button" name="b1" value=<%=sDeleteSymbol%> onclick="if(confirm('確定要<%=sDeleteSymbol%>此舉發單？')){DelBill_NoDCI(<%=trim(rsfound("SN"))%>)}" style="font-size: 10pt; width: 40px; height:26px;">	
		<%								else
		%>
								<input type="button" name="b1" value=<%=sDeleteSymbol%> onclick="if(confirm('確定要<%=sDeleteSymbol%>此舉發單？')){DelBill_NoDCI(<%=trim(rsfound("SN"))%>)}" style="font-size: 10pt; width: 40px; height:26px;">		
		<%								end if
									else	
									'已入案要寫DCILOG
		%>
								<input type="button" name="b1" value=<%=sDeleteSymbol%> onclick='window.open("BillBase_Del_DCI.asp?DBillSN=<%=trim(rsfound("SN"))%>","WebPage_Del_Bill","left=300,top=200,location=0,width=600,height=300,resizable=yes,scrollbars=yes")' style="font-size: 10pt; width: 40px; height:26px;">
		<%	
									end if
								end if
							else
		%>
								<input type="button" name="b1" value=<%=sDeleteSymbol%> onclick='window.open("BillBase_Del_DCI.asp?DBillSN=<%=trim(rsfound("SN"))%>","WebPage_Del_Bill","left=300,top=200,location=0,width=540,height=300,resizable=yes,scrollbars=yes")' style="font-size: 10pt; width: 40px; height:26px;">
		<%
							end if
						else	'PasserBase
								'1:查詢 ,2:新增 ,3:修改 ,4:刪除
								if CheckPermission(234,4)=true then
									if trim(rsfound("BillStatus"))="9" then
		%>
								<Input type="button" name="b1" value=<%=sDeleteSymbol%> onclick='alert("案件已結案，不可做刪除!!!!!");' style="font-size: 10pt; width: 40px; height:26px;">
		<%
									else
		%>
								<input type="button" name="b1" value=<%=sDeleteSymbol%> onclick='window.open("PasserBase_Del.asp?DBillSN=<%=trim(rsfound("SN"))%>","WebPage_Del_Passer","left=300,top=200,location=0,width=400,height=300,resizable=yes,scrollbars=yes")' style="font-size: 10pt; width: 40px; height:26px;">
		<%						end if
									end if 
						end if
					end if 
				end if
					'=======================================================
					if rsfound("BillStatus")="3" or rsfound("BillStatus")="4" or rsfound("BillStatus")="5" or rsfound("BillStatus")="7" or rsfound("BillStatus")="9" then
	%>					<input type="button" value="撤銷送達" onclick='window.open("../Query/BillSend_Dci_Cancel.asp?DBillSN=<%=trim(rsfound("SN"))%>","WebPage2_Cancel","left=100,top=100,location=0,width=400,height=250,resizable=yes,scrollbars=yes")' style="font-size: 10pt; width: 60px; height:26px;">
	<%
					end if
					if (sys_City="高雄市" Or sys_City=ApconfigureCityName) and trim(rsfound("BillTypeID"))="2" and trim(rsfound("BillBaseTypeID"))="0" then 
						%>						
						<input type="button" name="save" value="影像覆蓋" onclick='window.open("http://10.133.2.176/ReportImageUpload/Query/ImageRecover.asp?SN=<%=trim(rsfound("SN"))%>&UpMem=<%=Trim(session("Credit_ID"))%>","UploadReFile","left=0,top=0,location=0,width=900,height=665,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 60px; height:26px;">
						<!-- <input type="button" name="save100x" value="補傳民眾查詢影像" onclick='window.open("BillIllegalImageAdd.asp?SN=<%=trim(rsfound("SN"))%>","UploadReFile100x","left=0,top=0,location=0,width=900,height=665,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 120px; height:26px;"> -->
						<%	
					ElseIf (sys_City="台南市" or sys_City="保二總隊三大隊二中隊" or sys_City="彰化縣" or sys_City="南投縣" or sys_City="金門縣" or sys_City="基隆市") and trim(rsfound("BillBaseTypeID"))="0" then
						%>						
						<input type="button" name="save" value="影像覆蓋" onclick='window.open("ImageRecover.asp?SN=<%=trim(rsfound("SN"))%>&UpMem=<%=Trim(session("Credit_ID"))%>","UploadReFile","left=0,top=0,location=0,width=900,height=665,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 60px; height:26px;">
				
						<%	
					end if
					if sys_City="花蓮縣" Or sys_City="台東縣" then 
						if (Instr(rsfound("Rule1"),"56")>0) and (Instr(rsfound("Note"),"txt")>0) and isnull(rsfound("Billno")) then 
						%>						
						<input type="button" name="save" value="地址調整" onclick='window.open("AddressUpdate.asp?sys_CarNo=<%=trim(rsfound("CarNo"))%>&FileName=<%=trim(rsfound("ImageFileNameB"))%>&BillSN=<%=trim(rsfound("Sn"))%>","AddressUpdate","left=0,top=0,location=0,width=900,height=665,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt;  height:26px;">
						<%				
						end if
					end if
				else
					'已刪除之告發單則顯示出刪除原因
					DelNote=""
					DelReasonID=""
					if trim(rsfound("BillBaseTypeID"))="0" then
						strRea="select a.DelReason,a.Note,b.Content from BillDeleteReason a,DCIcode b where a.BillSN="&trim(rsfound("SN"))&" and b.TypeID=3 and b.ID=a.DelReason"
					else
						strRea="select a.DelReason,a.Note,b.Content from PasserDeleteReason a,DCIcode b where a.PasserSN="&trim(rsfound("SN"))&" and b.TypeID=3 and b.ID=a.DelReason"
					end if
					set rsRea=conn.execute(strRea)
					if not rsRea.eof then
						DelReasonID=trim(rsRea("DelReason"))
						DelNote=rsRea("Content")
						if trim(rsRea("Note"))<>"" and not isnull(rsRea("Note")) then
							DelNote=DelNote&"("&trim(rsRea("Note"))&")"
						end if
					end if
					rsRea.close
					set rsRea=nothing
					if (sys_City="高雄市" Or sys_City=ApconfigureCityName) and trim(rsfound("BillTypeID"))="2" and DelStatusID_tmp=1 then 
						if DelReasonID<>"Y" and DelReasonID<>"AAA" and DelReasonID<>"AAB" and DelReasonID<>"1" and DelReasonID<>"2" and DelReasonID<>"Z1" and DelReasonID<>"Z2" and DelReasonID<>"Z3"  then
						%>
						<input type="button" name="save" value="另案舉發" onclick='window.open("/traffic/BillKeyIn/BillKeyIn_Car_Report.asp?BillReCover=1&ReCoverSn=<%=trim(rsfound("SN"))%>","UploadFile","left=0,top=0,location=0,width=1010,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 60px; height:26px;">
						<%
						end if
					end if
					response.write DelNote
				end if
					'smith 20090714 台南市新增可上傳案件的申訴掃描檔等等 影像資料，後續於舉發單查詢可以做
					if sys_City="台南市" Or sys_City="基隆市" Or sys_City="花蓮縣" Or (sys_City="苗栗縣" And Trim(Session("Unit_ID"))="03BA") Or sys_City="澎湖縣" Or sys_City="金門縣" Or sys_City="保二總隊三大隊一中隊" Or sys_City="保二總隊第二大隊第一中隊" then 
						%>						
						<input type="button" name="save" value="上傳影像" onclick='window.open("UploadDetailFile.asp?SN=<%=trim(rsfound("SN"))%>","UploadFile","left=0,top=0,location=0,width=600,height=465,resizable=yes,scrollbars=yes,menubar=yes,status=yes")' style="font-size: 10pt; width: 60px; height:26px;">
						<%	
					ElseIf sys_City="高雄市XXX" Then
						%>						
						<input type="button" name="save" value="上傳影像" onclick='window.open("http://10.133.2.176/ReportImageUpload/Query/UploadDetailFile.asp?SN=<%=trim(rsfound("SN"))%>&UpMem=<%=Trim(session("Credit_ID"))%>","UploadFile","left=0,top=0,location=0,width=600,height=465,resizable=yes,scrollbars=yes,menubar=yes,status=yes")' style="font-size: 10pt; width: 60px; height:26px;">
						<%
					end If 
					
					If sys_City="苗栗縣" Then
					%>
						<input type="button" name="save" value="存查聯" onclick='window.open("BillBaseFastPaper_miaoli.asp?PBillSN=<%=trim(rsfound("SN"))%>","BillPrintPDF_miaoli","left=0,top=0,location=0,width=1010,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 60px; height:26px;">
					<%
					end if
					 
						response.write "</td>"
						response.write "</tr>"
						Response.flush
						rsfound.movenext
					Next
					rsfound.close
					Set rsfound=nothing
				end If
				
				
				%>

				
				
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			
			<input type="button" name="MovePage" value="跳至" onclick="funcPageGo();">  &nbsp;
			<input type="text" value="<%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)%>" name="DataPageNo" size="3" onkeyup="value=value.replace(/[^\d]/g,'')" >			
			頁
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">			
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
		<%	If (sys_City="苗栗縣" And Trim(Session("Unit_ID"))<>"03BA") Then

			else
		%>
			<input type="button" name="Submit1e32" value="攔停建檔清冊" onclick="funPrintCaseList_Stop();">
			<span class="style3"><img src="space.gif" width="13" height="8"></span>
			<input type="button" name="Submite32" value="逕舉建檔清冊" onclick="funPrintCaseList_Report();">
		<%
			End If 
			if sys_City<>"高雄市" and sys_City<>ApconfigureCityName then 
				If (sys_City="苗栗縣" And Trim(Session("Unit_ID"))<>"03BA") Then

				else
			%>
				<span class="style3"><img src="space.gif" width="13" height="8"></span>
				<input type="button" name="Submit4232" value="列印車籍清冊" onclick="funchgCarDataList();">
			<%
				End If 
			end if
			%>
			<%
			if sys_City="台東縣" then 
			%>
				<span class="style3"><img src="space.gif" width="13" height="8"></span>
				<input type="button" name="Submit4232" value="列印戶籍地址補正車籍清冊" onclick="funchgownerDataList();">

				<span class="style3"><img src="space.gif" width="13" height="8"></span>
				<input type="button" name="Submit4232" value="匯出戶籍地址補正清冊" onclick="funchgownerDataList2();">
			<%
			end if
			%>

			<span class="style3"><img src="space.gif" width="5" height="8"></span>
		<%
			If Trim(session("Credit_ID"))="A000000000" And sys_City="彰化縣" Then
		%>
			<input type="button" name="btnExecel" value="彰化審計室資料" onclick="funchgExecel_CH();">
		<%	
			End If

			If (sys_City="苗栗縣" And Trim(Session("Unit_ID"))<>"03BA") Then

			else
		%>
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<bR>
			<font size="2">EXCEL檔案格式,匯出限制不可大於6萬筆紀錄</font>
		<%	End if%>
			<input type="hidden" name="DelReason" value="">
		</td>
	</tr>

</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="Del_SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
		<%'response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
		<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&request("Sys_RecordMemberID")&"');"%>
	function funSelt(){
		var error=0;
		var errorString="";
	<%If sys_City="台南市" then%>
		if(myForm.Sys_BillNo.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：因資安審查規定，查詢必須輸入舉發單單號!!";
		}
	<%end if%>
		if(myForm.IllegalDate.value!=""){
			if(!dateCheck(myForm.IllegalDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}
		if(myForm.IllegalDate1.value!=""){
			if(!dateCheck(myForm.IllegalDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}
	
		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}

		if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.BillFillDate.value!=""){
			if(!dateCheck(myForm.BillFillDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：填單日期輸入不正確!!";
			}
		}
		
		if(myForm.QryReason.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：因資安審查規定，查詢必須選擇查詢事由!!";
		}
	<%if sys_City="台中市" or sys_City="高雄市" then%>
		if (myForm.Sys_BillNo.value=="" && myForm.Sys_CarNo.value==""){
			myForm.RecordDateCheck.checked=true;

			rdateTmp=myForm.RecordDate.value;
			ryear=parseInt(rdateTmp.substr(0,rdateTmp.length-4))+1911;
			rmonth=rdateTmp.substr(rdateTmp.length-4,2);
			rday=rdateTmp.substr(rdateTmp.length-2,2);
			var RecDate=new Date(ryear,rmonth-1,rday);
			var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
			var OverDate=new Date();
			OverDate=DateAdd("d",-1,DateAdd2("m",-3,thisDay));
			//alert(OverDate +"\n"+ RecDate);
			if (OverDate > RecDate){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期起始日不可大於三個月!!";
			}
		}
	<%else%>
		if(myForm.IllegalDateCheck.checked==false && myForm.RecordDateCheck.checked==false && myForm.BillFillDateCheck.checked==false && myForm.Sys_BillNo.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請勾選並輸入任一日期區間，如有單號可不勾選日期區間!!";
		}		
	<%end if%>
		/*
		if(myForm.RecordDate_h.value!="" || myForm.RecordDate1_h.value!=""){
			if(myForm.RecordDate_h.value=="" || myForm.RecordDate1_h.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔時段輸入不完整!!";
			}
		}
		*/
		if (error>0){
			alert(errorString);
		}else{
			myForm.btnSelt.disabled=true;
			myForm.btnSelt.value="查詢中請稍候";
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}
	function Selt_Stop(){
		var error=0;
		var errorString="";

		if (error>0){
			alert(errorString);
		}else{
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt_Stop";
			myForm.submit();
		}
	}

	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
	}
	function repage(){
		myForm.DB_Move.value=0;
		myForm.submit();
	}
	function funchgExecel(){
		UrlStr="PasserAndBillBaseQry_Execel.asp";
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_CH(){
		UrlStr="PasserAndBillBaseQry_Execel_CH.asp";
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}

	//列印車籍清冊
	function funchgCarDataList(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲列印車籍清冊的舉發單！");
		}else{
			UrlStr="PrintCarDataList.asp?QryReason="+myForm.QryReason.value;
			newWin(UrlStr,"CarListWin",900,575,0,0,"yes","yes","yes","no");
		}
	}

	//列印戶籍地址補正車籍清冊
	function funchgownerDataList(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲列印戶籍地址車籍清冊的舉發單！");
		}else{
			UrlStr="PrintOwnerDataList.asp";
			newWin(UrlStr,"CarListWin",900,575,0,0,"yes","yes","yes","no");
		}
	}

	//匯出戶籍地址補正清冊
	function funchgownerDataList2(){
		if (myForm.DB_Selt.value==""){
			alert("請先查詢欲匯出戶籍地址補正清冊的舉發單！");
		}else{
			UrlStr="PrintDriverAddrList.asp";
			newWin(UrlStr,"CarListWin",900,575,0,0,"yes","yes","yes","no");
		}
	}
	//攔停建檔清冊
	function funPrintCaseList_Stop(){
		UrlStr="PrintCaseDataList_Stop.asp";
		newWin(UrlStr,"CaseListWin",980,575,0,0,"yes","yes","yes","no");
	}
	//逕舉建檔清冊
	function funPrintCaseList_Report(){
		UrlStr="PrintCaseDataList_Report.asp";
		newWin(UrlStr,"CaseListWin",980,575,0,0,"yes","yes","yes","no");
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
	function DelBill_NoDCI(DelSN){
		myForm.Del_SN.value=DelSN;
		myForm.kinds.value="Del_NoDci";
		myForm.submit();
	}

function funcPageGo(){
	if (myForm.DataPageNo.value < 1 || myForm.DataPageNo.value > <%=fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%>){
		alert("頁數輸入錯誤!");
	}else{
		myForm.DB_Move.value=<%=(10+request("sys_MoveCnt"))%> * (myForm.DataPageNo.value-1);
		myForm.submit();
	}
}
<%if trim(request("Sys_RecordUnit"))="" and sys_City="高雄市" and trim(Session("Unit_ID"))="0807" then%>
myForm.Sys_RecordUnit.value="";
<%end if%>
</script>
<%
conn.close
set conn=nothing
%>