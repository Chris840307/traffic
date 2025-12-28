<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
If trim(Request("chk_City"))="台南市" and (not ifnull(Request("Ch_Name"))) Then

	strSQL="select * from memberdata where RecordstateID=0 and chName='"&trim(Request("Ch_Name"))&"' and UnitID in(select UnitID from Unitinfo where UnitName like '"&trim(Request("Unit_Name"))&"%') order by MODIFYTIME DESC"

	set rsmen=conn.execute(strSQL)

	If not rsmen.eof Then
		Session("FuncID")="233,1,1,1,1&&254,1,1,1,1&&252,1,1,1,1&&253,1,1,1,1&&250,1,1,1,1&&297,1,1,1,1&&249,1,1,1,1&&255,1,1,1,1&&245,1,1,1,1&&228,1,1,1,1&&226,1,1,1,1&&290,1,1,1,1&&230,1,1,1,1&&281,1,1,1,1&&280,1,1,1,1&&1601,1,1,1,1&&285,1,1,1,1&&1500,1,1,1,1&&235,1,1,1,1&&224,1,1,1,1&&299,1,1,1,1&&236,1,1,1,1&&293,1,1,1,1&&270,1,1,1,1&&229,1,1,1,1&&298,1,1,1,1&&263,1,1,1,1&&1729,1,1,1,1&&220,1,1,1,1&&260,1,1,1,1&&262,1,1,1,1&&261,1,1,1,1&&227,1,1,1,1&&901,1,1,1,1&&234,1,1,1,1&&223,1,1,1,1&&221,1,1,1,1"

		Session("Unit_ID")=rsmen("UnitID")
		Session("User_ID")=rsmen("MemberID")
		Session("Ch_Name")=rsmen("ChName")
		Session("Credit_ID")=rsmen("CreditID")
		Session("Group_ID")=rsmen("GroupRoleID")
		Session("ManagerPower")=rsmen("ManagerPower")
		'單位等級
		strUnit="select UnitLevelID,DCIwindowName from UnitInfo where UnitID='"&trim(rsmen("UnitID"))&"'"
		set rsUnit=conn.execute(strUnit)
		if not rsUnit.eof then
			UnitLevel=trim(rsUnit("UnitLevelID"))
			DCIwindow=trim(rsUnit("DCIwindowName"))
		end if
		rsUnit.close
		set rsUnit=nothing
		Session("UnitLevelID")=UnitLevel
		Session("DCIwindowName")=DCIwindow
	else
		Response.Redirect "/traffic/Traffic_Login.asp?Error=無使用之權限,請洽工程師。"
	End if
	rsmen.close
End If 

if isempty(request("DB_Selt")) then
	strSQL="update PasserBase set forfeit2=null where rule2 is null and forfeit2 is not null and recordDate between to_date(TO_CHAR(SYSDATE-14, 'YYYY/MM/DD')||' 00:00:00','YYYY/MM/DD/HH24/MI/SS') and to_date(TO_CHAR(SYSDATE, 'YYYY/MM/DD')||' 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

	conn.execute(strSQL)
end if

%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人道路障礙裁罰</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
.btn3{
   font-size:16px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
   font-weight:900;
}
</style>

</head>
<%
'檢查是否可進入本系統
AuthorityCheck(224)
DB_Orderby=split("Driver,IllegaLDate",",")
DB_KindSelt=trim(request("DB_KindSelt"))
DB_Selt=trim(request("DB_Selt"))
DB_Display=request("DB_Display")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false

if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then 

	showCreditor=true

	If isempty(request("DB_Selt")) Then

		'strSQL="delete passersenddetail where exists(select 'Y' from (select sendnumber,senddate from (select nvl(sendnumber,'1') sendnumber,nvl(senddate,to_date('1999/01/01','YYYY/MM/DD')) senddate,count(1) cnt from passersenddetail group by nvl(sendnumber,'1'),nvl(senddate,to_date('1999/01/01','YYYY/MM/DD')) ) tba where cnt >1 ) tab where sendnumber=nvl(passersenddetail.sendnumber,'1') and senddate=nvl(passersenddetail.senddate,to_date('1999/01/01','YYYY/MM/DD'))) and Not Exists(select 'N' from PASSERCREDITOR where senddetailsn=passersenddetail.SN)"

		'conn.execute(strSQL)
	
	End if 
	
end If

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

'==========================================================================================


'==========================================================================================

if DB_Selt="Selt" then
	strwhere="":strDate=""
	'日期
	if trim(request("Sys_DoubleCheckStatus"))<>"" then
		strwhere=strwhere&" and DoubleCheckStatus="&trim(request("Sys_DoubleCheckStatus"))
	end if

	if request("RecordDate1")<>"" and request("RecordDate2")<>""then
		ArgueDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.RecordDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("BillFillDate1")<>"" and request("BillFillDate2")<>""then
		ArgueDate1=gOutDT(request("BillFillDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("BillFillDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.BillFillDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("IllegalDate1")<>"" and request("IllegalDate2")<>""then
		ArgueDate1=gOutDT(request("IllegalDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("IllegalDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("UrgeDate1")<>"" and request("UrgeDate2")<>""then
		ArgueDate1=gOutDT(request("UrgeDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("UrgeDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(Select 'Y' from PasserUrge where billsn=a.sn and UrgeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS'))"
	end if

	if request("PayDate1")<>"" and request("PayDate2")<>""then
		ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(Select 'Y' from PasserPay where billsn=a.sn and PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS'))"
	end if

	if request("JudeDate1")<>"" and request("JudeDate2")<>""then
		ArgueDate1=gOutDT(request("JudeDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("JudeDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(Select 'Y' from PasserJude where billsn=a.sn and JudeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS'))"
	end if

	if request("SendDate1")<>"" and request("SendDate2")<>""then
		ArgueDate1=gOutDT(request("SendDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("SendDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(Select 'Y' from PasserSend where billsn=a.sn and SendDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS'))"

		If showCreditor Then

			strwhere=strwhere&" and (select count(1) cnt from PasserSendDetail where BillSn=a.SN)<=1"
		
		End if 

	end if

	if request("DeallIneDate1")<>"" and request("DeallIneDate2")<>"" then
		ArgueDate1=gOutDT(request("DeallIneDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("DeallIneDate2"))&" 23:59:59"
		strwhere=strwhere&" and a.DeallIneDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("Sys_SendDetailDate1")<>"" and request("Sys_SendDetailDate2")<>"" then
		ArgueDate1=gOutDT(request("Sys_SendDetailDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_SendDetailDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(select 'Y' from PasserSendDetail where SendDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and billsn=a.sn) and (select count(1) cnt from PasserSendDetail dt where billsn=a.sn)>1"
	end If 
	
	if request("Sys_CreditorType")<>"" then
		
		strwhere=strwhere&" and Not Exists(select 'Y' from PasserCreditor where billsn=a.sn)"
	end If

	if request("MakeSureDate1")<>"" and request("MakeSureDate2")<>""then
		ArgueDate1=gOutDT(request("MakeSureDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("MakeSureDate2"))&" 23:59:59"
		strwhere=strwhere&" and Exists(select 'Y' from PasserSend where MakeSureDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and billsn=a.sn)"
	end If 
	
	if request("CaseCloseDate1")<>"" and request("CaseCloseDate2")<>""then
		CaseCloseDate1=gOutDT(request("CaseCloseDate1"))&" 0:0:0"
		CaseCloseDate2=gOutDT(request("CaseCloseDate2"))&" 23:59:59"
		strwhere=strwhere&" and Exists(select 'Y' from PasserPay where CaseCloseDate between TO_DATE('"&CaseCloseDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&CaseCloseDate2&"','YYYY/MM/DD/HH24/MI/SS') and BillSN=a.SN)"
	end if

	if trim(request("sys_CreditorTypeID"))<>"" Or (trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"") Then
		strPasserCreditorAdd=""
		If trim(request("sys_CreditorTypeID")) <> "-1" Then
			If trim(request("sys_CreditorTypeID"))<>"" Then
				strPasserCreditorAdd=" CreditorTypeID in('"&trim(request("sys_CreditorTypeID"))&"')"
			End If 
			If trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"" Then
				PetitionDate1=gOutDT(request("Sys_PetitionDate1"))&" 0:0:0"
				PetitionDate2=gOutDT(request("Sys_PetitionDate2"))&" 23:59:59"
				If strPasserCreditorAdd="" Then
					strPasserCreditorAdd=" PetitionDate between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				Else
					strPasserCreditorAdd=strPasserCreditorAdd&" and PetitionDate between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				End If 
			End If
		end if

		If trim(request("sys_CreditorTypeID")) = "-1" Then
			strwhere=strwhere&" and not Exists(select 'N' from PasserSendDetail where Exists(select 'Y' from PasserCreditor where CreditorTypeID in('0','1') and SendDetailSN=PasserSendDetail.SN) and billsn=a.sn)"
		else
			strwhere=strwhere&" and Exists(select 'Y' from PasserSendDetail where Exists(select 'Y' from PasserCreditor where "&strPasserCreditorAdd&" and SendDetailSN=PasserSendDetail.SN) and BillSn=a.SN)"
		End if		
		
	end if 

	'單位
	if trim(request("Sys_BillUnitID"))<>"" then
		strwhere=strwhere&" and a.BillUnitID in ('"&request("Sys_BillUnitID")&"')"
	end if 

	'舉發人
	if request("Sys_BillMem")<>"" then
		strwhere=strwhere&" and (a.BillMem1='"&request("Sys_BillMem")&"' or a.BillMem2='"&request("Sys_BillMem")&"' or a.BillMem3='"&request("Sys_BillMem")&"')"
	end If 

	'建檔人
	'	if request("Sys_RecordMemberID")<>"" then
	'		strwhere=strwhere&" and a.RecordMemberID ="&request("Sys_RecordMemberID")
	'	end if
	'階段、繳費待確認

	'單號
	if request("Sys_BillNo")<>"" then
		strwhere=strwhere&" and a.BillNo ='" & Ucase(request("Sys_BillNo")) &  "'"
	end If 

	'違規人姓名
	if request("Sys_Driver")<>"" then
		strwhere=strwhere&" and a.Driver='"&request("Sys_Driver")&"'"
	end If 
 
	'違規人身分証號
	if request("Sys_DriverID")<>"" then
		strwhere=strwhere&" and a.DriverID='"&Ucase(request("Sys_DriverID"))&"'"
	end if

	if request("Sys_BillMemID")<>"" then
		strwhere=strwhere&" and a.BillMemID1 in(select MemberID from MemberData where LoginID='"&trim(request("Sys_BillMemID"))&"')"
	end if

	if request("Sys_BILLSTATUS")<>"" then
		if trim(request("Sys_BILLSTATUS"))="9" then
			strwhere=strwhere&" and a.BILLSTATUS=9"
		
		elseif trim(request("Sys_BILLSTATUS"))="1" then
			strwhere=strwhere&" and a.BILLSTATUS=9 and exists(select 'Y' from PasserPay where billsn=a.sn and PayAmount>0)"

		elseif trim(request("Sys_BILLSTATUS"))="2" then
			strwhere=strwhere&" and a.BILLSTATUS=9 and (select nvl(sum(PayAmount),0) from PasserPay where billsn=a.sn)=0"

		elseif trim(request("Sys_BILLSTATUS"))="3" then
			strwhere=strwhere&" and a.BILLSTATUS<>9 and exists(select 'Y' from PasserPay where billsn=a.sn and PayAmount>0)"

		elseif trim(request("Sys_BILLSTATUS"))="4" then
			strwhere=strwhere&"and exists(select 'Y' from PasserPay where billsn=a.sn)"

		else
			strwhere=strwhere&" and a.BILLSTATUS<>9"
		end if
	end If 

	if request("Sys_Rule")<>"" then
		strwhere=strwhere&" and (Rule1 like '"&trim(request("Sys_Rule"))&"%' or Rule2 like '"&trim(request("Sys_Rule"))&"%' or Rule3 like '"&trim(request("Sys_Rule"))&"%' or Rule4 like '"&trim(request("Sys_Rule"))&"%')"
	end If 
	
	if trim(request("Sys_Fastener1"))<>"" then
		strwhere=strwhere&" and exists(select 'Y' from PasserConfiscate where ConfiscateID='"&trim(request("Sys_Fastener1"))&"' and billsn=a.sn)"
	end if 

	if trim(request("Sys_SendNumber"))<>"" then
		strwhere=strwhere&" and exists(select 'Y' from PasserSend where SendNumber='"&trim(request("Sys_SendNumber"))&"' and billsn=a.sn)"
	end if 

	if trim(request("Sys_PayNo"))<>"" then
		strwhere=strwhere&" and exists(select 'Y' from passerpay where PayNo like '%"&trim(request("Sys_PayNo"))&"%' and billsn=a.sn)"
	end if 


	if trim(request("Sys_SendCase"))<>"" then
		if trim(request("Sys_SendCase"))="1" then
			strwhere=strwhere&" and e.ArrivedDate is null"
		elseif trim(request("Sys_SendCase"))="2" then
			strwhere=strwhere&" and Not(e.ArrivedDate is null)"
		end if
	end if

	'if request("Sys_BillTypeID")<>"" then
	'		strwhere=strwhere&" and a.BillTypeID="&request("Sys_BillTypeID")
	'end if

	if request("Sys_MemberStation")<>"" then
		strwhere=strwhere&" and a.MemberStation in('"&request("Sys_MemberStation")&"')"
	end If 
	
	if request("Sys_ReserveYear1")<>"" and request("Sys_ReserveYear2")<>"" then
		strwhere=strwhere&" and a.ReserveYear between "&request("Sys_ReserveYear1")&" and "&request("Sys_ReserveYear2")
	end if

'	if sys_City="基隆市" then
'		if DB_KindSelt="UrgeDateSelt" then
'			if trim(request("Sys_JudeDate"))<>"" then
'				ArgueDate1=DateAdd("d",0-Cint(request("Sys_JudeDate")),Date)&" 0:0:0"
'
'				strwhere=strwhere&" and a.DeallineDate < TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and a.BILLSTATUS<>9 and Not Exists(Select 'N' from PasserJude where billsn=a.sn)"
'
'			else
'
'				strwhere=strwhere&" and Not Exists(Select 'N' from PasserJude where billsn=a.sn) and a.BILLSTATUS<>9"
'			end if
'		end if
'
'		if DB_KindSelt="JudeDateSelt" then
'			if trim(request("Sys_UrgeDate"))<>"" then
'				ArgueDate1=DateAdd("d",0-Cint(request("Sys_UrgeDate")),Date)&" 0:0:0"
'
'				strwhere=strwhere&" and a.UrgeDate < TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and ((d.URGEDATE is null or d.URGEDATE=N'') and a.BILLSTATUS<>9"
'			else
'
'				strwhere=strwhere&" and (b.JUDEDATE is not null or b.JUDEDATE<>N'') and (d.URGEDATE is null or d.URGEDATE=N'') and a.BILLSTATUS<>9"
'			end if
'		end if
'
'		if DB_KindSelt="SendDateSelt" then
'			if trim(request("Sys_SendDate"))<>"" then
'				ArgueDate1=DateAdd("d",0-Cint(request("Sys_SendDate")),Date)&" 0:0:0"
'
'				strwhere=strwhere&" and b.JUDEDATE < TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and (c.SENDDATE is null or c.SENDDATE=N'') and a.BILLSTATUS<>9"
'			else
'
'				strwhere=strwhere&" and (b.JUDEDATE is not null or b.JUDEDATE<>N'') and (c.SENDDATE is null or c.SENDDATE=N'') and a.BILLSTATUS<>9"
'			end if
'		end if
'	else
'		if DB_KindSelt="UrgeDateSelt" then
'			strwhere=strwhere&" and b.JUDEDATE is null and d.URGEDATE is null and a.BILLSTATUS<>9"
'		end if
'
'		if DB_KindSelt="JudeDateSelt" then
'			strwhere=strwhere&" and b.JUDEDATE is not null and d.URGEDATE is null and c.SENDDATE is null and a.BILLSTATUS<>9"
'		end if
'
'		if DB_KindSelt="SendDateSelt" then
'			strwhere=strwhere&" and b.JUDEDATE is not null and (c.SENDDATE is null or c.SENDDATE=N'') and a.BILLSTATUS<>9"
'		end if
'	end If 
	
	if DB_KindSelt="UrgeDateSelt" then
		strwhere=strwhere&" and not Exists(select 'N' from PasserJude where BillSN=a.SN) and not Exists(select 'N' from PasserUrge where BillSN=a.SN) and a.BILLSTATUS<>9"
	end if

	if DB_KindSelt="JudeDateSelt" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where BillSN=a.SN) and not Exists(select 'N' from PasserUrge where BillSN=a.SN) and not Exists(select 'N' from PasserSend where BillSN=a.SN) and a.BILLSTATUS<>9"
	end if

	if DB_KindSelt="SendDateSelt" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where BillSN=a.SN) and not Exists(select 'N' from PasserSend where BillSN=a.SN) and a.BILLSTATUS<>9"
	end if

	if trim(request("Sys_Order"))<>"" then
		orderstr=" order by "&request("Sys_Order")
	end if
end if

if DB_Selt="SltUrgeDate" then
	'if request("Sys_sltUrgeDate1")<>"" and request("Sys_sltUrgeDate2")<>""then
	if request("Sys_sltUrgeDate1")<>"" then
		ArgueDate1=gOutDT(request("Sys_sltUrgeDate1"))&" 0:0:0"
		'ArgueDate2=gOutDT(request("Sys_sltUrgeDate2"))&" 23:59:59"
		ArgueDate2=gOutDT(request("Sys_sltUrgeDate1"))&" 23:59:59"
		strwhere=strwhere&" and d.UrgeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
end if

if DB_Selt="SltJudeDate" then
	'if request("Sys_sltJudeDate1")<>"" and request("Sys_sltJudeDate2")<>""then
	if request("Sys_sltJudeDate1")<>"" then
		ArgueDate1=gOutDT(request("Sys_sltJudeDate1"))&" 0:0:0"
		'ArgueDate2=gOutDT(request("Sys_sltJudeDate2"))&" 23:59:59"
		ArgueDate2=gOutDT(request("Sys_sltJudeDate1"))&" 23:59:59"
		strwhere=strwhere&" and b.JudeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
end if

if DB_Selt="SltSendDate" then
	'if request("Sys_sltSendDate1")<>"" and request("Sys_sltSendDate2")<>""then
	if request("Sys_sltSendDate1")<>"" then
		ArgueDate1=gOutDT(request("Sys_sltSendDate1"))&" 0:0:0"
		'ArgueDate2=gOutDT(request("Sys_sltSendDate2"))&" 23:59:59"
		ArgueDate2=gOutDT(request("Sys_sltSendDate1"))&" 23:59:59"
		strwhere=strwhere&" and c.SendDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
end if


if DB_Selt="Selt" then

	SQL_Save="delete PasserReportSave where UserID='"&Session("User_ID")&"' and ReportID='PasserBaseQry.asp'"
	conn.execute(SQL_Save)

	If trim(Request("Selt_SpanName")) <> "," and trim(Request("Selt_SpanName"))<>"" Then
		tmp_SpanName=split(Request("Selt_SpanName"),",")
		tmp_Span=split(Request("Selt_Span"),",")
		
		For i = 1 to Ubound(tmp_Span)-1
			tmp_ObjName=split(tmp_SpanName(i),"@")
			tmp_Obj=split(tmp_Span(i),"@")
			
			SQL_Save="insert into PasserReportSave(userid,reportid,typeid,L_xy,objid)" &_
			" values(" &_
			"'"&Session("User_ID")&"','PasserBaseQry.asp','Query'" &_
			",'"&tmp_ObjName(0)&"','"&tmp_ObjName(1)&"'" &_
			")"
			
			conn.execute(SQL_Save)

			SQL_Save="insert into PasserReportSave(userid,reportid,typeid,L_xy,objid)" &_
			" values(" &_
			"'"&Session("User_ID")&"','PasserBaseQry.asp','Query'" &_
			",'"&tmp_Obj(0)&"','"&tmp_Obj(1)&"'" &_
			")"
			
			conn.execute(SQL_Save)

		Next

	End if 

	If trim(Request("Selt_Rpt")) <> "," and trim(Request("Selt_Rpt"))<>"" Then
		tmp_Rpt=split(Request("Selt_Rpt"),",")
		
		For i = 1 to Ubound(tmp_Rpt)-1
			tmp_Obj=split(tmp_Rpt(i),"@")

			SQL_Save="insert into PasserReportSave(userid,reportid,typeid,L_xy,objid)" &_
			" values(" &_
			"'"&Session("User_ID")&"','PasserBaseQry.asp','Report'" &_
			",'"&tmp_Obj(0)&"','"&tmp_Obj(1)&"'" &_
			")"
			
			conn.execute(SQL_Save)

		Next

	End if 
	

	if trim(strwhere)="" then
		DB_Selt="":DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end If 

Sys_SendBillSN=trim(request("Sys_SendBillSN"))
if DB_Display="show" then

	showFiled=""
	If showCreditor Then
		showFiled=",(select max(PetitionDate) PetitionDate from PasserCreditor where billsn=a.SN) PetitionDate"
	End if 	
	
	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.IllegalAddress,a.Rule1," &_
	"a.RuleVer,a.FORFEIT1,a.FORFEIT2,a.FORFEIT3,a.FORFEIT4,a.BILLSTATUS," &_
	"a.BillMem1,a.DoubleCheckStatus," &_
	"(Select JudeDate from PasserJude where billsn=a.sn) JUDEDATE," &_
	"(Select SendDate from PasserSend where billsn=a.sn) SENDDATE," &_
	"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
	"(Select UrgeDate from PasserUrge where billsn=a.sn) URGEDATE," &_
	"(select Max(ArrivedDate) ArrivedDate from PassersEndArrived where PasserSN=a.sn) ArrivedDate," &_
	"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate"&showFiled &_
	" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	'set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere

	strSQL="select a.SN,a.BillNo,a.DoubleCheckStatus,a.Driver,a.IllegalDate,a.Rule1"&showFiled&" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	
	cntUbound=0
	set rssn=conn.execute(strSQL)
	BillSN="":BillNo="":DoubleCheckStatus=""
	while Not rssn.eof
		cntUbound=cntUbound+1
		if trim(BillSN)<>"" then
			BillSN=BillSN&","
			BillNo=BillNo&","
		end if
		BillSN=BillSN&rssn("SN")	
		BillNo=BillNo&rssn("BillNo")
		rssn.movenext
	wend
	rssn.close

	strCnt="select count(1) as cnt from PasserBase a where RecordStateID=0 "&strwhere
	set Dbrs=conn.execute(strCnt)
	DBsum=0
	If Not Dbrs.eof Then
		DBsum=cint(Dbrs("cnt"))
	end if
	Dbrs.close

	if sys_City="宜蘭縣" or sys_City="澎湖縣" then

		strSQL="select 'Y' from PasserBase a where RecordStateID=0 "&strwhere

		upSQL="Update PasserBase set ForFeit1=(select max(Level2) from law where version="&RuleVer&" and ItemID=PasserBase.Rule1),ForFeit2=(select max(Level2) from law where version="&RuleVer&" and ItemID=PasserBase.Rule2) where Exists("&strSQL&"  and a.SN=PasserBase.Sn) and DeallineDate < to_Date('"&date&"','YYYY/MM/DD') and BillStatus=0 and RecordStateid=0 and Exists(select 'Y' from law where itemid=PasserBase.Rule1 and Level1=PasserBase.ForFeit1 and VerSion="&RuleVer&")"

		conn.execute(upSQL)
	end If 	
	'response.write strSQLTemp
	'response.end
end If 

If showCreditor Then

'	if sys_City="台中市" and isempty(request("DB_Selt")) then
'
'		errmsg="":PasserCnt=0:PasserSendCnt=0:MeberStation=""
'
'		If Session("UnitLevelID") > 1 then MemberStation=" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"
'
'		strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and Exists(select 'Y' from PasserSend where TRUNC(sysdate-SendDate) between 300 and 365 and Not exists(select 'N' from PasserCreditor where SENDDETAILSN in(select sn from PasserSendDetail where billsn=PasserSend.Billsn and SendDate=PasserSend.SendDate) and billsn=PasserBase.sn) and billsn=PasserBase.sn)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"
'
'		set rssn=conn.execute(strSQL)
'
'		while Not rssn.eof
'			PasserCnt=cdbl(rssn("cnt"))
'
'			if trim(PasserCnt)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserCnt & "筆已移送後逾10個月且於一年內尚未取得債權\n"
'
'			rssn.movenext
'		wend
'		
'		rssn.close
'
'		If not ifnull(errmsg) Then
'			Response.write "<script>"
'			Response.Write "alert('" & errmsg & "！');"
'			Response.write "</script>"
'		end if
'
'	end if

	if (sys_City="基隆市" or sys_City="屏東縣") and isempty(request("DB_Selt")) then
		
		errmsg="":PasserCnt=0:PasserSendCnt=0:MeberStation=""

		If Session("UnitLevelID") > 1 then MemberStation=" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"

		strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and TRUNC(sysdate-DeallineDate)>60 and BillFillDate > to_date('"&(year(now)-6)&"/12/31','YYYY/MM/DD') and not exists(select 'Y' from PasserJude where billsn=PasserBase.SN) and Not exists(select 'Y' from PasserSend where billsn=PasserBase.SN)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"

		set rssn=conn.execute(strSQL)

		while Not rssn.eof
			PasserCnt=cdbl(rssn("cnt"))

			if trim(PasserCnt)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserCnt & "筆已逾60天尚未裁決\n"

			rssn.movenext
		wend
		
		rssn.close

		strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and BillFillDate > to_date('"&(year(now)-6)&"/12/31','YYYY/MM/DD') and exists(select 'Y' from PasserJude where billsn=PasserBase.SN and TRUNC(sysdate-JudeDate)>365) and Not exists(select 'Y' from PasserSend where billsn=PasserBase.SN)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"

		set rssn=conn.execute(strSQL)

		while Not rssn.eof
			PasserSendCnt=cdbl(rssn("cnt"))

			if trim(PasserSendCnt)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserSendCnt & "筆已裁決逾一年未移送\n"

			rssn.movenext
		wend
		
		rssn.close
		
		If not ifnull(errmsg) Then
			Response.write "<script>"
			Response.Write "alert('" & errmsg & "！');"
			Response.write "</script>"
		end if
		
	end If 
End if 

chk_SpanName=",QuyName01,QuyName02,QuyName03,QuyName04,QuyName05," &_
				"QuyName06,QuyName07,QuyName08,QuyName09,QuyName10," &_
				"QuyName11,QuyName12,QuyName13,QuyName14,QuyName15," &_
				"QuyName16,QuyName17,QuyName18,QuyName19,QuyName20," &_
				"QuyName21,QuyName22,QuyName23,QuyName24,"

chk_Span=",QuyObj01,QuyObj02,QuyObj03,QuyObj04,QuyObj05," &_
			"QuyObj06,QuyObj07,QuyObj08,QuyObj09,QuyObj10," &_
			"QuyObj11,QuyObj12,QuyObj13,QuyObj14,QuyObj15," &_
			"QuyObj16,QuyObj17,QuyObj18,QuyObj19,QuyObj20," &_
			"QuyObj21,QuyObj22,QuyObj23,QuyObj24,"

chk_Rpt=",rptObj01,rptObj02,rptObj03,rptObj04,rptObj05," &_
		"rptObj06,rptObj07,rptObj08,rptObj09,rptObj10," &_
		"rptObj11,rptObj12,rptObj13,rptObj14,rptObj15," &_
		"rptObj16,rptObj17,rptObj18,rptObj19,rptObj20," &_
		"rptObj21,rptObj22,rptObj23,rptObj24,rptObj25," &_
		"rptObj26,rptObj27,rptObj28,rptObj29,rptObj30," &_
		"rptObj31,rptObj32,rptObj33,rptObj34,"

Selt_SpanName=","
Selt_Span=","
Selt_Rpt=","
js_obj=""

obj_SQL="select userid,reportid,typeid,L_xy,objid" &_
		" from PasserReportSave" &_
		" where UserID='"&Session("User_ID")&"' and ReportID='PasserBaseQry.asp'" &_
		" order by typeid,L_xy"

set rsobj=conn.execute(obj_SQL)

While not rsobj.eof
	
	js_obj=js_obj & rsobj("L_xy") & ".innerHTML=" & rsobj("objid")&";"& vblf
	
	if Instr(rsobj("objid"),"_name")>0 then 

		Selt_SpanName=Selt_SpanName & rsobj("L_xy") & "@" & rsobj("objid") & ","
	else

		If Trim(rsobj("typeid")) = "Query" Then

			Selt_Span=Selt_Span & rsobj("L_xy") & "@" & rsobj("objid") & ","

			js_obj=js_obj&"for(i=0;i<=myForm.seleObj.length;i++){"& vblf
			js_obj=js_obj&"	if(myForm.seleObj[i].value=='"&rsobj("objid")&"'){"& vblf
			js_obj=js_obj&"		myForm.seleObj[i].checked=true;"& vblf
			js_obj=js_obj&"		break;"& vblf
			js_obj=js_obj&"	}"& vblf
			js_obj=js_obj&"}"& vblf
		else

			js_obj=js_obj&"for(i=0;i<=myForm.RptSelet.length;i++){"& vblf
			js_obj=js_obj&"	if(myForm.RptSelet[i].value=='"&rsobj("objid")&"'){"& vblf
			js_obj=js_obj&"		myForm.RptSelet[i].checked=true;"& vblf
			js_obj=js_obj&"		break;"& vblf
			js_obj=js_obj&"	}"& vblf
			js_obj=js_obj&"}"& vblf

			Selt_Rpt=Selt_Rpt & rsobj("L_xy") & "@" & rsobj("objid") & ","
		end If 

	end If 

	rsobj.movenext
Wend
rsobj.close

%>
<body onLoad="funLoadSend();">
<form name="myForm" method="post">
<div id="menu" style="position:absolute; visibility:hidden;">
	<table width="50" border="0" cellspacing="1" cellpadding="0" onContextMenu="return false;" style="cursor:hand">
		<tr>
			<td bgcolor="#BBBBBB">
				<table width="100%" border="0" cellspacing="1" cellpadding="1">
					<tr bgcolor="#EEEEEE">
						<td colspan="3" align="center" style="height:30px;font-size:18px;line-height: 25px">
							<b>查&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;詢&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;條&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;件&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;設&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;定</b>
						</td>
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_DoubleCheckStatus" onclick="Check_Selt(this,'Sys_DoubleCheckStatus_name','Sys_DoubleCheckStatus');">
							建檔序號

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_BillUnitID" onclick="Check_Selt(this,'Sys_BillUnitID_name','Sys_BillUnitID');">
							舉發單位

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_BILLSTATUS" onclick="Check_Selt(this,'Sys_BILLSTATUS_name','Sys_BILLSTATUS');">
							繳費狀況

							&nbsp;&nbsp;
						</td>
						
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_Driver" onclick="Check_Selt(this,'Sys_Driver_name','Sys_Driver');">
							違規人名

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_DriverID" onclick="Check_Selt(this,'Sys_DriverID_name','Sys_DriverID');">
							身分證號

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_BillMemID" onclick="Check_Selt(this,'Sys_BillMemID_name','Sys_BillMemID');">
							舉發人代碼

							&nbsp;&nbsp;
						</td>
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="IllegalDate1" onclick="Check_Selt(this,'IllegalDate1_name','IllegalDate1');">
							違規日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="UrgeDate1" onclick="Check_Selt(this,'UrgeDate1_name','UrgeDate1');">
							催告日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="PayDate1" onclick="Check_Selt(this,'PayDate1_name','PayDate1');">
							付費日期

							&nbsp;&nbsp;
						</td>
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="JudeDate1" onclick="Check_Selt(this,'JudeDate1_name','JudeDate1');">
							裁決日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="SendDate1" onclick="Check_Selt(this,'SendDate1_name','SendDate1');">
							移送日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="DeallIneDate1" onclick="Check_Selt(this,'DeallIneDate1_name','DeallIneDate1');">
							應到案日期

							&nbsp;&nbsp;
						</td>
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="RecordDate1" onclick="Check_Selt(this,'RecordDate1_name','RecordDate1');">
							建檔日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="BillFillDate1" onclick="Check_Selt(this,'BillFillDate1_name','BillFillDate1');">
							填單日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_Rule" onclick="Check_Selt(this,'Sys_Rule_name','Sys_Rule');">
							法條代碼

							&nbsp;&nbsp;
						</td>
					</tr>
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="CaseCloseDate1" onclick="Check_Selt(this,'CaseCloseDate1_name','CaseCloseDate1');">
							結案日期

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_Fastener1" onclick="Check_Selt(this,'Sys_Fastener1_name','Sys_Fastener1');">
							代保管物

							&nbsp;&nbsp;
						</td>
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="seleObj" value="Sys_PayNo" onclick="Check_Selt(this,'Sys_PayNo_name','Sys_PayNo');">
							收據號碼

							&nbsp;&nbsp;
						</td>
					</tr>
					<%
					If showCreditor then
					%>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="Sys_SendDetailDate1" onclick="Check_Selt(this,'Sys_SendDetailDate1_name','Sys_SendDetailDate1');">
								再次移送日

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="sys_CreditorTypeID" onclick="Check_Selt(this,'sys_CreditorTypeID_name','sys_CreditorTypeID');">
								債權狀況

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="MakeSureDate1" onclick="Check_Selt(this,'MakeSureDate1_name','MakeSureDate1');">
								確定日期

								&nbsp;&nbsp;
							</td>
						</tr>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="Sys_PetitionDate1" onclick="Check_Selt(this,'Sys_PetitionDate1_name','Sys_PetitionDate1');">
								債權取得日

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="Sys_ReserveYear" onclick="Check_Selt(this,'Sys_ReserveYear_name','Sys_ReserveYear');">
								保留年度

								&nbsp;&nbsp;
							</td>
							
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="seleObj" value="Sys_SendNumber" onclick="Check_Selt(this,'Sys_SendNumber_name','Sys_SendNumber');">
								移送號案

								&nbsp;&nbsp;
							</td>
						</tr>
					<%
					end If 
					%>

					<tr bgcolor="#EEEEEE">
						<td colspan="3" align="center" style="height:30px;font-size:18px;line-height: 25px">
							<b>報&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;表&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;設&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;定</b>
						</td>
					</tr>

					
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_CrtList" onclick="Check_Rpt(this,'btn_CrtList');">
							建檔清冊

							&nbsp;&nbsp;
						</td>
					
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_illegalExl" onclick="Check_Rpt(this,'btn_illegalExl');">
							舉發清冊(xls)

							&nbsp;&nbsp;
						</td>
					
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_JudeList" onclick="Check_Rpt(this,'btn_JudeList');">
							裁決清冊(xls)

							&nbsp;&nbsp;
						</td>
					</tr>
					
					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_JudeOpenList" onclick="Check_Rpt(this,'btn_JudeOpenList');">
							裁決公示送達清冊

							&nbsp;&nbsp;
						</td>
					
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_JudeSendList" onclick="Check_Rpt(this,'btn_JudeSendList');">
							裁決寄存送達清冊

							&nbsp;&nbsp;
						</td>
					
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_SendDetailList" onclick="Check_Rpt(this,'btn_SendDetailList');">
							移送案件明細表

							&nbsp;&nbsp;
						</td>
					</tr>

					<tr bgcolor="#EEEEEE">

						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayDetailList" onclick="Check_Rpt(this,'btn_PayDetailList');">
							繳費明細表

							&nbsp;&nbsp;
						</td>

						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReport" onclick="Check_Rpt(this,'btn_PayReport');">
							收繳費統計表

							&nbsp;&nbsp;
						</td>

						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_NoPayDetailList" onclick="Check_Rpt(this,'btn_NoPayDetailList');">
							未繳費明細表

							&nbsp;&nbsp;
						</td>
					</tr>

					<tr bgcolor="#EEEEEE">
						<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_NoPayReport" onclick="Check_Rpt(this,'btn_NoPayReport');">
							未繳費統計表

							&nbsp;&nbsp;
						</td>

						<td colspan="2" onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
							&nbsp;&nbsp;

							<input class="btn1" type="checkbox" name="RptSelet" value="btn_YearSaveList" onclick="Check_Rpt(this,'btn_YearSaveList');">
							年度保留清冊

							&nbsp;&nbsp;
						</td>
					</tr>
					<%
					if sys_City<>"基隆市" then 
					%>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_UrgeList" onclick="Check_Rpt(this,'btn_UrgeList');">
								催告清冊

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_UrgeCloseList" onclick="Check_Rpt(this,'btn_UrgeCloseList');">
								催告已到案清冊

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_UrgeNotCloseList" onclick="Check_Rpt(this,'btn_UrgeNotCloseList');">
								催告未到案清冊

								&nbsp;&nbsp;
							</td>
						</tr>
					<%
					end if
					%>

					<%
					if showCreditor then 
					%>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_NotArrivedList" onclick="Check_Rpt(this,'btn_NotArrivedList');">
								未送達清冊

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_NoPayExeList" onclick="Check_Rpt(this,'btn_NoPayExeList');">
								未繳納待執行清冊

								&nbsp;&nbsp;
							</td>							

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_SendnotCreditList" onclick="Check_Rpt(this,'btn_SendnotCreditList');">
								已移送未執行債權清冊

								&nbsp;&nbsp;
							</td>
						</tr>

						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_SendtwoList" onclick="Check_Rpt(this,'btn_SendtwoList');">
								債權憑證準備再移送清冊

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_InventoryList" onclick="Check_Rpt(this,'btn_InventoryList');">
								交付保管品核對清冊

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_NoEffectsList" onclick="Check_Rpt(this,'btn_NoEffectsList');">
								無個人財產清冊

								&nbsp;&nbsp;
							</td>
						</tr>

						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_CreditorList" onclick="Check_Rpt(this,'btn_CreditorList');">
								債權憑證清冊

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_CreditorStstusList" onclick="Check_Rpt(this,'btn_CreditorStstusList');">
								債權執行狀態清冊

								&nbsp;&nbsp;
							</td>
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_SendCreditorList" onclick="Check_Rpt(this,'btn_SendCreditorList');">
								債權明細統計表

								&nbsp;&nbsp;
							</td>
						</tr>

						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_SendCreditorTwoList" onclick="Check_Rpt(this,'btn_SendCreditorTwoList');">
								債權再移送明細統計表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayDetail_YiLanList" onclick="Check_Rpt(this,'btn_PayDetail_YiLanList');">
								行政罰鍰收繳情形明細表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_NoPayDetail_YiLanList" onclick="Check_Rpt(this,'btn_NoPayDetail_YiLanList');">
								應收未收收繳情形明細表

								&nbsp;&nbsp;
							</td>
						</tr>
					<%
					end if
					%>

					<%
					if sys_City="彰化縣" then
					%>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayCityReport" onclick="Check_Rpt(this,'btn_PayCityReport');">
								(署)交通管理事件統計表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReportMonth" onclick="Check_Rpt(this,'btn_PayReportMonth');">
								交通罰緩收入憑證月報表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReportMonthUnit" onclick="Check_Rpt(this,'btn_PayReportMonthUnit');">
								(分局)交通罰緩收入憑證月報表

								&nbsp;&nbsp;
							</td>
						</tr>

						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReportDayUnit" onclick="Check_Rpt(this,'btn_PayReportDayUnit');">
								(分局)交通罰緩收據明細表

								&nbsp;&nbsp;
							</td>
						</tr>

					<%
					end if
					%>

					<%
					if sys_City="台中市" then
					%>
						<tr bgcolor="#EEEEEE">
							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReport2" onclick="Check_Rpt(this,'btn_PayReport2');">
								各單位清理統計表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReport3" onclick="Check_Rpt(this,'btn_PayReport3');">
								各單位清理進度表

								&nbsp;&nbsp;
							</td>

							<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
								&nbsp;&nbsp;

								<input class="btn1" type="checkbox" name="RptSelet" value="btn_PayReport4" onclick="Check_Rpt(this,'btn_PayReport4');">
								交通違規罰鍰繳款明細表

								&nbsp;&nbsp;
							</td>
						</tr>

					<%
					end if
					%>

					<tr bgcolor="#EEEEEE" onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';">
						<td colspan="3" align="center" style="height:30px;font-size:18px;line-height: 25px" onClick="AddQuery()" title="隱藏此功能表不做任何動作">
							<b>確&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;定</b>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>
<div id="RptOpton" style="position:absolute;left:0px; top:200px;">
	<table border="0" cellspacing="0" cellpadding="0" onContextMenu="return false;" style="cursor:hand">
		<tr bgcolor="#ffffff">
			<td valign="top">
				<div id="prt_menu" style="position:absolute; visibility:hidden;width:300px;">
					<table width="100%" border="0" cellspacing="1" cellpadding="1">
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj01" name="rptObj01" onclick="ChangeRpt('rptObj01');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj02" name="rptObj02" onclick="ChangeRpt('rptObj02');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj03" name="rptObj03" onclick="ChangeRpt('rptObj03');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj04" name="rptObj04" onclick="ChangeRpt('rptObj04');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj05" name="rptObj05" onclick="ChangeRpt('rptObj05');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj06" name="rptObj06" onclick="ChangeRpt('rptObj06');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj07" name="rptObj07" onclick="ChangeRpt('rptObj07');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj08" name="rptObj08" onclick="ChangeRpt('rptObj08');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj09" name="rptObj09" onclick="ChangeRpt('rptObj09');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj10" name="rptObj10" onclick="ChangeRpt('rptObj10');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj11" name="rptObj11" onclick="ChangeRpt('rptObj11');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj12" name="rptObj12" onclick="ChangeRpt('rptObj12');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj13" name="rptObj13" onclick="ChangeRpt('rptObj13');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj14" name="rptObj14" onclick="ChangeRpt('rptObj14');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj15" name="rptObj15" onclick="ChangeRpt('rptObj15');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj16" name="rptObj16" onclick="ChangeRpt('rptObj16');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj17" name="rptObj17" onclick="ChangeRpt('rptObj17');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj18" name="rptObj18" onclick="ChangeRpt('rptObj18');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj19" name="rptObj19" onclick="ChangeRpt('rptObj19');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj20" name="rptObj20" onclick="ChangeRpt('rptObj20');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj21" name="rptObj21" onclick="ChangeRpt('rptObj21');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj22" name="rptObj22" onclick="ChangeRpt('rptObj22');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj23" name="rptObj23" onclick="ChangeRpt('rptObj23');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj24" name="rptObj24" onclick="ChangeRpt('rptObj24');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj25" name="rptObj25" onclick="ChangeRpt('rptObj25');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj26" name="rptObj26" onclick="ChangeRpt('rptObj26');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj27" name="rptObj27" onclick="ChangeRpt('rptObj27');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj28" name="rptObj28" onclick="ChangeRpt('rptObj28');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj29" name="rptObj29" onclick="ChangeRpt('rptObj29');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj30" name="rptObj30" onclick="ChangeRpt('rptObj30');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj31" name="rptObj31" onclick="ChangeRpt('rptObj31');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj32" name="rptObj32" onclick="ChangeRpt('rptObj32');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj33" name="rptObj33" onclick="ChangeRpt('rptObj33');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
						<tr align="right" bgcolor="#ffffff">
							<td>
								<span id="rptObj34" name="rptObj34" onclick="ChangeRpt('rptObj34');" style="background-color:#ffffff;"></span>
							</td>
						</tr>
					</table>
				</div>
			</td>
			<td style="font-size:16px;width:15px;background-color:#BBBBBB;" onclick="AddReportMenu();">
				<b>
					<br><br><br>報<br><br><br>表<br><br><br>選<br><br><br>單<br><br><br>
				</b>
			</td>
		</tr>
	</table>
</div>

<table width="100%" border="0">
	<tr height="30">
		<td height="30" bgcolor="#FFCC33">
		<font size="3"><b>慢車行人道路障礙舉發單紀錄</b> </font><img src="space.gif" width="32" height="10">
		<a href="passerbase.doc" target="_blank" ><font size="3"> 下載 裁罰系統使用說明.doc</font></a></img>
		<%if showCreditor then%>
			<B>
			<a href="PasserCreditor.doc" target="_blank" ><font size="3" color="red"> 下載 債權憑證系統使用說明.doc</font></a>
			<a href="PasserCreditorReport.doc" target="_blank" ><font size="3" color="red"> 下載 債權憑證報表使用說明.doc</font></a>
			</td>
			</b>
			
		<%end if%>
	</tr>
	<tr>
		<td bgcolor="#cccccc">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#ffffff">
				<tr>
					<td>
						<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
							<tr>
								<td style="line-height:20px;">
									<strong><font color="red">舉發單號</font></strong>
								</td>
								<td>
									<input name="Sys_BillNo" maxlength="9" size="8" class="btn1" type="text" value="<%=Ucase(request("Sys_BillNo"))%>" size="8" maxlength="20">
								</td>
								<td style="line-height:20px;" nowrap>
									應到案處
								</td>
								<td style="line-height:20px;">
									<select name="Sys_MemberStation" class="btn1"><%
										strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
										set rsUnit=conn.execute(strSQL)
										If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
										rsUnit.close
										strUnitID="":strtmp=""
										if trim(Session("UnitLevelID"))="1" then
											strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
											strtmp=strtmp+"<option value="""">所有單位</option>"
										else
											strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID=(select UnitTypeid from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"

											set rs1=conn.execute(strSQL)
											while Not rs1.eof
												if trim(strUnitID)="" and rs1("UnitLevelID")>1 then
													strUnitID=trim(rs1("UnitID"))
												elseif rs1("UnitLevelID")>1 then 
													strUnitID=strUnitID&"','"&trim(rs1("UnitID"))
												end if
												rs1.movenext
											wend
											rs1.close
											strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
										end if
										set rs1=conn.execute(strSQL)
										while Not rs1.eof
											strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
											if trim(rs1("UnitID"))=trim(request("Sys_MemberStation")) then
												strtmp=strtmp+" selected"
											end if
											strtmp=strtmp+">"&rs1("UnitName")&"</option>"
											rs1.movenext
										wend
										rs1.close
										strtmp=strtmp+"</select>"
										response.write strtmp%>
								</td>
								<td style="line-height:20px;">
									列表排序
								</td>
								<td style="line-height:20px;">
									<select Name="Sys_Order" class="btn1">
										<option value="DoubleCheckStatus,BillNo"<%if trim(request("Sys_Order"))="DoubleCheckStatus,BillNo" then response.write " selected"%>>建檔序號</option>
										<option value="Driver,BillNo"<%if trim(request("Sys_Order"))="Driver,BillNo" then response.write " selected"%>>違規人</option>
										<option value="IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="IllegaLDate,BillNo" then response.write " selected"%>>違規日期</option>
										<option value="Rule1,Driver,IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="Rule1,Driver,IllegaLDate,BillNo" then response.write " selected"%>>法條,違規人,違規日期</option>
										<option value="Driver,IllegaLDate,Rule1,BillNo"<%if trim(request("Sys_Order"))="Driver,IllegaLDate,Rule1,BillNo" then response.write " selected"%>>違規人,違規日期,法條</option>
										<%if showCreditor then%>
											<option value="PetitionDate,Driver,BillNo"<%if trim(request("Sys_Order"))="PetitionDate,Driver,BillNo" then response.write " selected"%>>取得債權日,違規人,單號</option>
										<%end if%>
										
									</select>
								</td>
							</tr>
							<tr>
								<td style="line-height:20px;">
									<span id="QuyName01" name="QuyName01" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName01');"></span>

								</td>
								<td style="line-height:20px;">
									<span id="QuyObj01" name="QuyObj01" style="background-color:#ffffff;"></span>

								</td>
								<td style="line-height:20px;">									
									<span id="QuyName02" name="QuyName02" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName02');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj02" name="QuyObj02" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName03" name="QuyName03" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName03');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj03" name="QuyObj03" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							<tr>
								<td style="line-height:20px;">									
									<span id="QuyName04" name="QuyName04" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName04');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj04" name="QuyObj04" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName05" name="QuyName05" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName05');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj05" name="QuyObj05" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName06" name="QuyName06" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName06');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj06" name="QuyObj06" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							<tr>
								<td style="line-height:20px;">									
									<span id="QuyName07" name="QuyName07" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName07');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj07" name="QuyObj07" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName08" name="QuyName08" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName08');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj08" name="QuyObj08" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName09" name="QuyName09" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName09');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj09" name="QuyObj09" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							<tr>
								<td style="line-height:20px;">									
									<span id="QuyName10" name="QuyName10" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName10');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj10" name="QuyObj10" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName11" name="QuyName11" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName11');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj11" name="QuyObj11" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName12" name="QuyName12" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName12');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj12" name="QuyObj12" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							<tr>
								<td style="line-height:20px;">									
									<span id="QuyName13" name="QuyName13" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName13');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj13" name="QuyObj13" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName14" name="QuyName14" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName14');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj14" name="QuyObj14" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName15" name="QuyName15" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName15');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj15" name="QuyObj15" style="background-color:#ffffff;"></span>								
								</td>
							</tr>						
							<tr>
								<td style="line-height:20px;">									
									<span id="QuyName16" name="QuyName16" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName16');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj16" name="QuyObj16" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName17" name="QuyName17" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName17');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj17" name="QuyObj17" style="background-color:#ffffff;"></span>								
								</td>
							</tr>

							<tr bgcolor="#CCFFCC">
								<td style="line-height:20px;">									
									<span id="QuyName18" name="QuyName18" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName18');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj18" name="QuyObj18" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName19" name="QuyName19" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName19');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj19" name="QuyObj19" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName20" name="QuyName20" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName20');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj20" name="QuyObj20" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							<tr bgcolor="#CCFFCC">
								<td style="line-height:20px;">									
									<span id="QuyName21" name="QuyName21" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName21');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj21" name="QuyObj21" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName22" name="QuyName22" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName22');"></span>
									
								</td>
								<td style="line-height:20px;" nowrap>
									<span id="QuyObj22" name="QuyObj22" style="background-color:#ffffff;"></span>								
								</td>
								<td style="line-height:20px;">									
									<span id="QuyName23" name="QuyName23" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName23');"></span>
									
								</td>
								<td style="line-height:20px;" nowrap>
									<span id="QuyObj23" name="QuyObj23" style="background-color:#ffffff;"></span>								
								</td>

							</tr>
							<tr>
																
								<td style="line-height:20px;">									
									<span id="QuyName24" name="QuyName24" style="background-color:#ffffff;" onclick="ChangeLocation('QuyName24');"></span>
									
								</td>
								<td style="line-height:20px;">									
									<span id="QuyObj24" name="QuyObj24" style="background-color:#ffffff;"></span>								
								</td>
							</tr>
							
							<tr>
								<td style="line-height:20px;" colspan="6" align="right">
									<input type="button" name="btnSelt" value="═╬═" class="btn3" style="width:70px;height:25px;" onclick="AddQuery();">
									<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funSelt();">
									<input type="button" name="btnCls" value="清除" class="btn3" style="width:60px;height:25px;" onClick="location='PasserBaseQry_1070817.asp'">
									<input type="button" name="btnSelt" value="══" class="btn3" style="width:50px;height:25px;" onclick="DelQuery();">
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									
								</td>
							</tr>
						</table>						
						<HR>
							未做裁決且未繳費
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funUrgeDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							裁決後
							未繳費且未催告
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funJudeDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							裁決後
							未繳費且未移送
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funSendDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							<input type="button" name="btnSelt" class="btn3" style="width:90px;height:25px;" value="資料回復" onclick="funRecallData();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,3)=false then
								response.write "disabled"
							end if
							%>>
							<%'if sys_City="基隆市" or sys_City="高雄市" or sys_City="台中市" then%>
							&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
							<input type="button" name="btnSelt" class="btn3" style="width:170px;height:25px;" value="整批戶籍地址更正" onclick="funUpdAddress();" >
							<%'end if%>
					</td>
				</tr>
			</table>
		</td> 
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">
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
			筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong>
			<font color="red"><B>操作項目若不勾選，即為全部處理。</B></font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10" nowrap>操作<br>項目</th>
					<th class="font10" nowrap>序號</th>
					<th class="font10" nowrap>違規日</th>
					<th class="font10" nowrap>舉發單號</th>
					<th class="font10" nowrap>舉發人</th>
					<th class="font10" nowrap>違規人</th>
					<th class="font10" nowrap>法條</th>
					<th class="font10" nowrap>裁決日</th>
					<%if sys_City<>"基隆市" then%>
						<th class="font10" nowrap>催告日</th>
					<%end if%>
					<th class="font10" nowrap>移送日</th>
					<%if showCreditor then%>
						<th class="font10" nowrap>再移送日</th>
						<th class="font10" nowrap>債權日(1)</th>						
						<th class="font10" nowrap>債權日(N)</th>
					<%end if%>
					<th class="font10" nowrap>送達日</th>
					<%if showCreditor then%>
						<th class="font10" nowrap>確定日</th>
					<%end if%>
					<th class="font10" nowrap>繳費日</th>
					<th class="font10" nowrap>已繳<br>金額</th>
					<th class="font10" nowrap>結案<br>狀態</th>
					<th class="font10">操作 ( 單筆送達證書 請點選 送達 按鈕 )</th>
				</tr>
				<%
			if DB_Display="show" then
				set rsfound=conn.execute(strSQLTemp)
				if Trim(request("DB_Move"))="" then
					DBcnt=0
				else
					DBcnt=request("DB_Move")
				end if
				if Not rsfound.eof then rsfound.move DBcnt
				for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
					if rsfound.eof then exit for
					response.write "<tr align='center' bgcolor='#ffffff'"
					lightbarstyle 0
					response.write ">"
					response.write "<td class=""font10""><input class=""btn1"" type=""checkbox"" name=""chkSend"" value="""&trim(rsfound("Sn"))&""" onclick=funChkSend();></td>"
					response.write "<td class=""font10"">"&trim(rsfound("DoubleCheckStatus"))&"</td>"
					response.write "<td class=""font10"">"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("BillMem1"))&"</td>"
					'response.write "<td><a href='../BillKeyIn/BillKeyIn_People_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&trim(rsfound("BillNo"))&"</a></td>"

					response.write "<td class=""font10"">"&trim(rsfound("Driver"))&"</td>"
					'response.write "<td>"&trim(rsfound("IllegalAddress"))&"</td>"

					if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
					
					'strSQL="Select * from Law where ItemID='"&rsfound("Rule1")&"' and VerSion='"&rsfound("RuleVer")&"'"
					'set rslaw=conn.execute(strSQL)
					'if not rslaw.eof then
					'	if trim(rslaw("Level1"))<>"" then FORFEIT=trim(rslaw("Level1"))
					'	if trim(rslaw("Level2"))<>"" then FORFEIT=FORFEIT&","&trim(rslaw("Level2"))
					'	if trim(rslaw("Level3"))<>"" then FORFEIT=FORFEIT&","&trim(rslaw("Level3"))
					'	if trim(rslaw("Level4"))<>"" then FORFEIT=FORFEIT&","&trim(rslaw("Level4"))
					'end if
					'rslaw.close
					response.write "<td class=""font10"">"
					response.write chRule
					if trim(FORFEIT)<>"" then response.write "<font size=1>("&FORFEIT&")</font>"
					response.write "</td>"
					'response.write "<td>"&FORFEIT&"</td>"
					'smith 20091005 基隆不用秀
					response.write "<td class=""font10"">"&trim(gInitDT(rsfound("JUDEDATE")))&"</td>"

					if sys_City<>"基隆市" then
						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("URGEDATE")))&"</td>"
					End if 

					SENDDATEtmp=""
					if showCreditor then
						strSD="select min(SendDate) as SendDate from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and SendDate is not null"
						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	
							SENDDATEtmp=Trim(rsSD("SendDate"))
						End If
						rsSD.close
						Set rsSD=Nothing 
					end if

					If SENDDATEtmp="" Or IsNull(SENDDATEtmp) Then
						strSD2="select * from PasserSend where BillSn="&Trim(rsfound("sn"))
						Set rsSD2=conn.execute(strSD2)
						If Not rsSD2.eof Then
							SENDDATEtmp=Trim(rsSD2("SendDate"))
						End If
						rsSD2.close
						Set rsSD2=Nothing 
					End If 
					response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"

					if showCreditor then

						strSD="select max(SendDate) as SendDate from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and SendDate is not null"
						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	

							If SENDDATEtmp=Trim(rsSD("SendDate")) Then

								SENDDATEtmp=""
							else

								SENDDATEtmp=Trim(rsSD("SendDate"))
							End if 
						else

							SENDDATEtmp=""
						End If
						rsSD.close
						Set rsSD=Nothing 
						response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"
					end if
					

					'-------------------LEO修改------標頭有改，但是列表沒改------------------------------
					'response.write "<td class=""font10"">"&trim(gInitDT(rsfound("ArrivedDate")))&"</td>"

					if showCreditor then
						SENDDATEtmp=""

						strSD="select PetitionDate from PasserCreditor where senddetailsn=(select min(sn) from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and sendDate=(select min(sendDate) from PasserSendDetail mins where BillSn="&Trim(rsfound("sn"))&"))"

						Set rsSD=conn.execute(strSD)

						If Not rsSD.eof then	
							SENDDATEtmp=Trim(rsSD("PetitionDate"))
						End If
						rsSD.close
						Set rsSD=Nothing 

						response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"

						SENDDATEtmp2=""

						strSD="select PetitionDate from PasserCreditor where senddetailsn=(select max(sn) from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and sendDate=(select max(sendDate) from PasserSendDetail mins where BillSn="&Trim(rsfound("sn"))&"))"

						Set rsSD=conn.execute(strSD)

						If Not rsSD.eof then	
							SENDDATEtmp2=Trim(rsSD("PetitionDate"))

							If SENDDATEtmp = SENDDATEtmp2 Then SENDDATEtmp2=""
						End If
						rsSD.close
						Set rsSD=Nothing 

						response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp2))&"</td>"


						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("ArrivedDate")))&"</td>"
						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("MakeSureDate")))&"</td>"
					Else
						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("ArrivedDate")))&"</td>"
					End if
					'------------------------------------------------------------------------------------------
					response.write "<td class=""font10"">"&trim(gInitDT(rsfound("PayDate")))&"</td>"

					Sys_Payamount=0
					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
					set rspay=conn.execute(strSQL)
					if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
					rspay.close

					response.write "<td class=""font10"">"&Sys_Payamount&"</td>"
					response.write "<td class=""font10"">"
					if trim(rsfound("BILLSTATUS"))="9" then
						response.write "結案."
					end if
					response.write "</td>"
%>					<td class="font10" nowrap align="left">

						<input type='button' value='繳款' class="btn3" style="width:50px;height:25px;" onclick='window.open("<%
							if sys_City="彰化縣" then

								Response.Write "Passer_Pay_Sys.asp?PBillSN="&trim(rsfound("SN"))
							else

								Response.Write "Passer_Pay.asp?PBillSN="&trim(rsfound("SN"))
							End if 
						%>","WebPage4","left=0,top=0,location=0,width=1100,height=575,resizable=yes,scrollbars=yes")'>						
					
						<!--<input type='button' value='執行處回文' onclick='window.open("Passer_Send.asp","WebPage2","left=0,top=0,location=0,width=500,height=455,resizable=yes,scrollbars=yes")'>		
						<br>-->

						<%if showCreditor then%>
							<input type='button' value='債權' class="btn3" style="width:50px;height:25px;" onclick='window.open("PasserCreditor.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=900,height=575,resizable=yes,scrollbars=yes")'<%If Ifnull(rsfound("SENDDATE")) Then Response.Write " disabled"%>>
						<%end if%>

						<input type="button" name="Update" value="詳細" class="btn3" style="width:50px;height:25px;" onclick='window.open("../Query/ViewBillBaseData_people.asp?BillSn=<%=trim(rsfound("SN"))%>","WebPage1","left=0,top=0,location=0,width=850,height=700,status=yes,resizable=yes,scrollbars=yes")'>

						<input type="button" name="btnSelt" value="其它" class="btn3" style="width:50px;height:25px;" onclick="ExePross('<%=trim(rsfound("SN"))%>');">

						<div id="exe_menu_<%=trim(rsfound("SN"))%>" style="position:absolute; visibility:hidden;">
							<table width="50" border="0" cellspacing="1" cellpadding="0" onContextMenu="return false;" style="cursor:hand">
								<tr>
									<td bgcolor="#BBBBBB">
										<table width="100%" border="0" cellspacing="1" cellpadding="1">
											<tr bgcolor="#EEEEEE">
												<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
													
													
													<input type='button' value='裁決' class="btn3" style="width:50px;height:25px;" onclick='exe_menu_<%=trim(rsfound("SN"))%>.style.visibility="hidden";window.open("Passer_Jude.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage1","left=0,top=0,location=0,width=950,height=700,status=yes,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>

												</td>
											</tr>
											<tr bgcolor="#EEEEEE">
												<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>
													
													<input type='button' value='送達' class="btn3" style="width:50px;height:25px;" onclick='exe_menu_<%=trim(rsfound("SN"))%>.style.visibility="hidden";window.open("PasserSendArrived.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=900,height=575,resizable=yes,scrollbars=yes")'>
												</td>
											</tr>
											
											<%
											if sys_City<> "基隆市" then
											%>

												<tr bgcolor="#EEEEEE">
													<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>

														<input type='button' value='催告' class="btn3" style="width:50px;height:25px;" onclick='exe_menu_<%=trim(rsfound("SN"))%>.style.visibility="hidden";window.open("PasserUrgeDetail.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage2","left=0,top=0,location=0,width=1000,height=600,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>
													</td>
												</tr>
												
											<%
											end If 
											%>
											
											<tr bgcolor="#EEEEEE">
												<td onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';" nowrap>

													<input type='button' value='移送' class="btn3" style="width:50px;height:25px;" onclick='exe_menu_<%=trim(rsfound("SN"))%>.style.visibility="hidden";window.open("PasserSendDetail.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage3","left=0,top=0,location=0,width=1000,height=600,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>
												</td>
											</tr>

											<tr bgcolor="#EEEEEE" onMouseOver="this.style.backgroundColor='#cccccc';" onMouseOut="this.style.backgroundColor='#EEEEEE';">
												<td colspan="3" align="center" style="height:30px;font-size:18px;line-height: 25px">

													<input type='button' value='取&nbsp;&nbsp;消' class="btn3" style="width:50px;height:25px;" onClick="ExePross('<%=trim(rsfound("SN"))%>')" title="隱藏此功能表不做任何動作">
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</div>
					</td>
<%
					response.write "</tr>"
				rsfound.MoveNext
				next
				rsfound.close
				set rsfound=nothing
			end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFDD77" align="center" nowrap>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);" class="btn3" style="width:60px;height:30px;font-size:14px;">
			<span class="style2"><%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);" class="btn3" style="width:60px;height:30px;font-size:14px;">
			<img src="space.gif" width="18" height="8">

			<input type="button" name="btnExecel" value="轉換成Excel" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funchgExecel();">
			<input type="button" name="btnExecel" value="郵局大宗函件" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funPasserMailMoney();">
			<br>
			
			<input type="button" name="btnExecel" value="批次裁決通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeBat();">

			<% '基隆催告的都不用秀 smith 20091005 
			if sys_City<>"基隆市" then %>
				<input type="button" name="btnExecel" value="批次催繳通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funUrgeBat();">
			<% end if %>

			<input type="button" name="btnExecel" value="批次移送通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funSendBat();">

			<input type="button" name="btnExecel" value="行政執行移送電子清冊" class="btn3" style="width:160px;height:30px;font-size:14px;" onclick="funCountryBat();">
			<%if showCreditor then%>
				<input type="button" name="btnExecel" class="btn3" style="width:120px;height:30px;font-size:14px;" value="批次債權移送" onclick="funSendBatTwo_chromat();">
			<%end if%>

			<!--
			<%if sys_City="台中市" or sys_City="台中縣" then%>
			<br>
			<input type="button" name="btnExecel" value="批次催繳套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funUrgeBat_chromat();">
			
			<input type="button" name="btnExecel" value="批次裁決套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeBat_chromat();">

			<input type="button" name="btnExecel" value="批次移送套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funSendBat_chromat();">
			
			<%end if%>
			-->
			
		</td>
	</tr>
	
	<tr>		
		<td height="35" bgcolor="#ffffff" align="center" nowrap>
			<!-- <font size="2">依據上方選擇案件資料產生相關清冊</font> -->

			<!--<input type="button" name="btnExecel" value="強制執行移送清冊" onclick="funUrgeJudeExecel();">-->

		<!--
		<br>
			<input type="button" name="btnprintBill" value="列印違規通知單" onclick="funPrintStyle()">
			<input type="button" name="btnprintBill" value="列印送達證書" onclick="funUrgeList()">
			 , 違規通知單與送達證書為Legal格式
		-->
		<br>
		<center><font size="2">清冊列印，皆為A4 直式格式</font></center>
		</td>
	</tr>

</table>
<input type="Hidden" name="chkSend" value="">
<input type="Hidden" name="DB_Selt" value="<%=trim(DB_Selt)%>">
<input type="Hidden" name="DB_KindSelt" value="<%=trim(DB_KindSelt)%>">
<input type="Hidden" name="DB_Display" value="<%=trim(DB_Display)%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="FromILLEGALDATE" value="<%=trim(request("ILLEGALDATE"))%>">
<input type="Hidden" name="TOILLEGALDATE" value="<%=trim(request("ILLEGALDATE1"))%>">
<input type="Hidden" name="orderstr" value="<%=orderstr%>">
<input type="Hidden" name="printStyle" value="">
<input type="Hidden" name="Sys_PasserNotify" value="">
<input type="Hidden" name="Sys_PasserSign" value="">
<input type="Hidden" name="Sys_PasserJude" value="">
<input type="Hidden" name="Sys_PasserJude_Label" value="">
<input type="Hidden" name="Sys_PasserLabel_miaoli" value="">
<input type="Hidden" name="Sys_PasserDeliver" value="">
<input type="Hidden" name="Sys_PasserSend" value="">
<input type="Hidden" name="Sys_PasserJudeSend" value="">
<input type="Hidden" name="Sys_PasserUrge" value="">

<input type="Hidden" name="MailMoneyValue" value="">
<input type="Hidden" name="Session_JudeName" value="">
<input type="Hidden" name="BillUrge" value="">
<input type="Hidden" name="Sys_SendBillSN" value="<%=Sys_SendBillSN%>">
<input type="Hidden" name="hd_BillSN" value="<%=BillSN%>">
<input type="Hidden" name="hd_BillNo" value="<%=BillNo%>">

<input type="Hidden" name="chk_SpanName" value="<%=chk_SpanName%>">
<input type="Hidden" name="chk_Span" value="<%=chk_Span%>">
<input type="Hidden" name="Selt_SpanName" value=<%=Selt_SpanName%>>
<input type="Hidden" name="Selt_Span" value="<%=Selt_Span%>">
<input type="Hidden" name="chk_Rpt" value="<%=chk_Rpt%>">
<input type="Hidden" name="Selt_Rpt" value="<%=Selt_Rpt%>">

<input type="Hidden" name="chk_button_event" value="">

<input type="Hidden" name="chk_button_Other" value="">

</form>
</body>
</html>

<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var PasserWin;
var strBillSN=myForm.Sys_SendBillSN.value;
var ck_Sn=strBillSN.split(',');
<%'response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
var Sys_City="<%=sys_City%>";

var tx="\u0022";


var Sys_DoubleCheckStatus_name="建檔序號";
var Sys_DoubleCheckStatus="<input name="+tx+"Sys_DoubleCheckStatus"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_DoubleCheckStatus")%>"+tx+" size="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">";


var Sys_BillUnitID_name="舉發單位";
var Sys_BillUnitID="<select name="+tx+"Sys_BillUnitID"+tx+" class="+tx+"btn1"+tx+"><%
						strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
						set rsUnit=conn.execute(strSQL)
						If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
						rsUnit.close
						strUnitID="":strtmp=""
						if trim(Session("UnitLevelID"))="1" then
							strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
							strtmp=strtmp+"<option value=""+tx+tx+"">所有單位</option>"
						elseif trim(Session("UnitLevelID"))="2" then
							strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID=(select UnitTypeid from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"

							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								if trim(strUnitID)="" and rs1("UnitLevelID")>1 then
									strUnitID=trim(rs1("UnitID"))
								elseif rs1("UnitLevelID")>1 then 
									strUnitID=strUnitID&"','"&trim(rs1("UnitID"))
								end if
								rs1.movenext
							wend
							rs1.close
							strtmp=strtmp+"<option value=""+tx+tx+"">所有單位</option>"
							'strtmp=strtmp+"<option value="""&strUnitID&""">管轄單位</option>"
						elseif trim(Session("UnitLevelID"))="3" then
							strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitTypeID,UnitName"
						end if
						set rs1=conn.execute(strSQL)
						while Not rs1.eof
							strtmp=strtmp+"<option value=""+tx+"""&rs1("UnitID")&"""+tx+"""
							if trim(rs1("UnitID"))=trim(request("Sys_BillUnitID")) then
								strtmp=strtmp+" selected"
							end if
							strtmp=strtmp+">"&rs1("UnitName")&"</option>"
							rs1.movenext
						wend
						rs1.close
						strtmp=strtmp+"</select>"
						response.write strtmp%>";


var Sys_BILLSTATUS_name="繳費狀況";
var Sys_BILLSTATUS="<select Name="+tx+"Sys_BILLSTATUS"+tx+" class="+tx+"btn1"+tx+">"+
						"<option value="+tx+tx+">全部</option>"+
						"<option value="+tx+"0"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="0" then response.write " selected"
						%>>未繳費</option>"+
						"<option value="+tx+"4"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="4" then response.write " selected"
						%>>已繳費(全部)</option>"+
						"<option value="+tx+"9"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="9" then response.write " selected"
						%>>已繳費(已結案含免罰)</option>"+
						"<option value="+tx+"1"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="1" then response.write " selected"
						%>>已繳費(已結案不含免罰)</option>"+
						"<option value="+tx+"2"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="2" then response.write " selected"
						%>>已繳費(免罰)</option>"+
						"<option value="+tx+"3"+tx+"<%
							if trim(request("Sys_BILLSTATUS"))="3" then response.write " selected"
						%>>已繳費(未結案)</option>"+
					"</select>";


var Sys_Driver_name="違規人名";
var Sys_Driver="<input name="+tx+"Sys_Driver"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_Driver")%>"+tx+
					" size="+tx+"7"+tx+" maxlength="+tx+"8"+tx+" onkeyup="+tx+"funSearchCname('Sys_Driver','SearChName');"+tx+" onMouseDown="+tx+"funCrtVale('Sys_Driver','SearChName','');"+tx+">"+
					"<br><div id="+tx+"SearChName"+tx+" style="+tx+"position:absolute;"+tx+"></div>";


var Sys_DriverID_name="身份證號";
var Sys_DriverID="<input name="+tx+"Sys_DriverID"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+" value="+tx+"<%=Ucase(request("Sys_DriverID"))%>"+tx+" size="+tx+"10"+tx+" maxlength="+tx+"12"+tx+">";


var Sys_BillMemID_name="舉發人代碼";
var Sys_BillMemID="<input name="+tx+"Sys_BillMemID"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_BillMemID")%>"+tx+
					" size="+tx+"7"+tx+" maxlength="+tx+"8"+tx+">";


var IllegalDate1_name="違規日期";
var IllegalDate1="<input name="+tx+"IllegalDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("IllegalDate1")%>"+tx+" size="+tx+"5"+tx+
						" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('IllegalDate1');"+tx+">"+
						"		~		"+
					"<input name="+tx+"IllegalDate2"+tx+" class="+tx+"btn1"+tx+
						" type="+tx+"text"+tx+" value="+tx+"<%=request("IllegalDate2")%>"+tx+
						" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('IllegalDate2');"+tx+">";


var IllegalDate1_name="違規日期";
var IllegalDate1="<input name="+tx+"IllegalDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("IllegalDate1")%>"+tx+" size="+tx+"5"+tx+
						" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('IllegalDate1');"+tx+">"+
						"		~		"+
					"<input name="+tx+"IllegalDate2"+tx+" class="+tx+"btn1"+tx+
						" type="+tx+"text"+tx+" value="+tx+"<%=request("IllegalDate2")%>"+tx+
						" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('IllegalDate2');"+tx+">";


var UrgeDate1_name="催告日期";
var UrgeDate1="<input name="+tx+"UrgeDate1"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("UrgeDate1")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('UrgeDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"UrgeDate2"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("UrgeDate2")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('UrgeDate2');"+tx+">";


var PayDate1_name="付費日期";
var PayDate1="<input name="+tx+"PayDate1"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("PayDate1")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('PayDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"PayDate2"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("PayDate2")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('PayDate2');"+tx+">";


var JudeDate1_name="裁決日期";
var JudeDate1="<input name="+tx+"JudeDate1"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("JudeDate1")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('JudeDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"JudeDate2"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("JudeDate2")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('JudeDate2');"+tx+">";

	
var SendDate1_name="移送日期";
var SendDate1="<input name="+tx+"SendDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("SendDate1")%>"+tx+" size="+tx+"5"+tx+
					" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('SendDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"SendDate2"+tx+" class="+tx+"btn1"+tx+
					" type="+tx+"text"+tx+" value="+tx+"<%=request("SendDate2")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('SendDate2');"+tx+">";



var DeallIneDate1_name="應到案日期";
var DeallIneDate1="<input name="+tx+"DeallIneDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("DeallIneDate1")%>"+tx+" size="+tx+"5"+tx+
						" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('DeallIneDate1');"+tx+">"+
						"		~		"+
						"<input name="+tx+"DeallIneDate2"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("DeallIneDate2")%>"+tx+
						" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('DeallIneDate2');"+tx+">";


var RecordDate1_name="建檔日期";
var RecordDate1="<input name="+tx+"RecordDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("RecordDate1")%>"+tx+" size="+tx+"5"+tx+
						" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
						" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('RecordDate1');"+tx+">"+
						"		~		"+
						"<input name="+tx+"RecordDate2"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("RecordDate2")%>"+tx+" size="+tx+"5"+tx+
						" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
						" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('RecordDate2');"+tx+">";


var BillFillDate1_name="填單日期";
var BillFillDate1="<input name="+tx+"BillFillDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("BillFillDate1")%>"+tx+" size="+tx+"5"+tx+
					" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
					" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('BillFillDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"BillFillDate2"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("BillFillDate2")%>"+tx+" size="+tx+"5"+tx+
					" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
					" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('BillFillDate2');"+tx+">";


var Sys_Rule_name="法條代碼";
var Sys_Rule="<input name="+tx+"Sys_Rule"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("Sys_Rule")%>"+tx+" size="+tx+"9"+tx+">";


var CaseCloseDate1_name="結案日期";
var CaseCloseDate1="<input name="+tx+"CaseCloseDate1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("CaseCloseDate1")%>"+tx+
					" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
					" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
					" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('CaseCloseDate1');"+tx+">"+
					"			~			"+
					"<input name="+tx+"CaseCloseDate2"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
					" value="+tx+"<%=request("CaseCloseDate2")%>"+tx+" size="+tx+"5"+tx+
					" maxlength="+tx+"8"+tx+" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
					" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:20px;height:20px;"+tx+
					" onclick="+tx+"OpenWindow('CaseCloseDate2');"+tx+">";

var Sys_SendNumber_name="移送號案";
var Sys_SendNumber="<input name="+tx+"Sys_SendNumber"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
				" value="+tx+"<%=request("Sys_SendNumber")%>"+tx+" size="+tx+"20"+tx+">";

var Sys_PayNo_name="收據號碼";
var Sys_PayNo="<input name="+tx+"Sys_PayNo"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
				" value="+tx+"<%=request("Sys_PayNo")%>"+tx+" size="+tx+"20"+tx+">";

var Sys_Fastener1_name="代保管物";
var Sys_Fastener1="<select Name="+tx+"Sys_Fastener1"+tx+" class="+tx+"btn1"+tx+">"+
					"<option value="+tx+tx+">全部</option><%
						strItem="select * from Code where TypeID=2 and Not(ID<478 or ID=479) order by ID"

						set rsItem=conn.execute(strItem)
						While Not rsItem.Eof
							Response.Write "<option value=""+tx+"""&trim(rsItem("ID"))&"""+tx+"""

							if trim(request("Sys_Fastener1"))=trim(rsItem("ID")) then response.write " selected"

							Response.Write ">"
							Response.Write trim(rsItem("Content"))
							Response.Write "</option>"
							rsItem.MoveNext
						Wend
						rsItem.close
						set rsItem=nothing
					%></select>";


var Sys_SendDetailDate1_name="再次移送日";
var Sys_SendDetailDate1="<input name="+tx+"Sys_SendDetailDate1"+tx+" class="+tx+"btn1"+tx+
							" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_SendDetailDate1")%>"+tx+
							" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
							" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
							"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
							" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
							" style="+tx+"width:20px;height:20px;"+tx+
							" onclick="+tx+"OpenWindow('Sys_SendDetailDate1');"+tx+">"+
							"			~			"+
							"<input name="+tx+"Sys_SendDetailDate2"+tx+" class="+tx+"btn1"+tx+
							" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_SendDetailDate2")%>"+tx+
							" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
							" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
							"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+" value="+tx+"..."+tx+
							" class="+tx+"btn3"+tx+" style="+tx+"width:20px;height:20px;"+tx+
							" onclick="+tx+"OpenWindow('Sys_SendDetailDate2');"+tx+">";


var sys_CreditorTypeID_name="債權狀況";
var sys_CreditorTypeID="<select name="+tx+"sys_CreditorTypeID"+tx+">"+
							"<option value="+tx+"-1"+tx+"<%
								if trim(Request("sys_CreditorTypeID"))="-1" then response.write " Selected"
							%>>未申請債權</option>"+
							"<option value="+tx+tx+"<%
								if trim(Request("sys_CreditorTypeID"))="" then response.write " Selected"
							%>>全部</option>"+
							"<option value="+tx+"0','1"+tx+"<%
								if trim(Request("sys_CreditorTypeID"))="0','1" then response.write " Selected"
							%>>已申請債權</option>"+
							"<option value="+tx+"0"+tx+"<%
								if trim(Request("sys_CreditorTypeID"))="0" then response.write " Selected"
							%>>清償中</option>"+
							"<option value="+tx+"1"+tx+"<%
								if trim(Request("sys_CreditorTypeID"))="1" then response.write " Selected"
							%>>無個人財產</option>"+
							"</select>";


var MakeSureDate1_name="確定日期";
var MakeSureDate1="<input name="+tx+"MakeSureDate1"+tx+" class="+tx+"btn1"+tx+
							" type="+tx+"text"+tx+" value="+tx+"<%=request("MakeSureDate1")%>"+tx+
							" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
							" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
							"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
							" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
							" style="+tx+"width:20px;height:20px;"+tx+
							" onclick="+tx+"OpenWindow('MakeSureDate1');"+tx+">"+
							"			~			"+
							"<input name="+tx+"MakeSureDate2"+tx+" class="+tx+"btn1"+tx+
							" type="+tx+"text"+tx+" value="+tx+"<%=request("MakeSureDate2")%>"+tx+
							" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
							" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
							"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
							" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
							" style="+tx+"width:20px;height:20px;"+tx+
							" onclick="+tx+"OpenWindow('MakeSureDate2');"+tx+">";


var Sys_PetitionDate1_name="債權取得日";
var Sys_PetitionDate1="<input name="+tx+"Sys_PetitionDate1"+tx+" class="+tx+"btn1"+tx+
						" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_PetitionDate1")%>"+tx+
						" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('Sys_PetitionDate1');"+tx+">"+
						"			~			"+
						"<input name="+tx+"Sys_PetitionDate2"+tx+" class="+tx+"btn1"+tx+
						" type="+tx+"text"+tx+" value="+tx+"<%=request("Sys_PetitionDate2")%>"+tx+
						" size="+tx+"5"+tx+" maxlength="+tx+"8"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
						"<input type="+tx+"button"+tx+" name="+tx+"datestr"+tx+
						" value="+tx+"..."+tx+" class="+tx+"btn3"+tx+
						" style="+tx+"width:20px;height:20px;"+tx+
						" onclick="+tx+"OpenWindow('Sys_PetitionDate2');"+tx+">";


var Sys_ReserveYear_name="保留年度";
var Sys_ReserveYear="<input name="+tx+"Sys_ReserveYear1"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("Sys_ReserveYear1")%>"+tx+" size="+tx+"2"+tx+" maxlength="+tx+"3"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">"+
					"&nbsp;~&nbsp;"+
					"<input name="+tx+"Sys_ReserveYear2"+tx+" class="+tx+"btn1"+tx+" type="+tx+"text"+tx+
						" value="+tx+"<%=request("Sys_ReserveYear2")%>"+tx+" size="+tx+"2"+" maxlength="+tx+"3"+tx+
						" onkeyup="+tx+"value=value.replace(/[^0-9\]/g, '');"+tx+">";

//===========================================================================================

var btn_CrtList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_CrtList"+tx+
					" value="+tx+"建檔清冊"+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
					" onclick="+tx+"myForm.chk_button_event.value=1;funPrintPeopleList_Stop();"+tx+">"+
					"&nbsp;&nbsp;";


var btn_illegalExl="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_illegalExl"+tx+
					" value="+tx+"舉發清冊(xls)"+tx+" class="+tx+"btn3"+tx+
					" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
					" onclick="+tx+"myForm.chk_button_event.value=1;funJudeExecel();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_JudeList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_JudeList"+tx+
				" value="+tx+"裁決清冊(xls)"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funJudeListExecel()"+tx+">"+
					"&nbsp;&nbsp;";

var btn_JudeOpenList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_JudeOpenList"+tx+
				" value="+tx+"裁決公示送達清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funGovArriveBat();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_JudeSendList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_JudeSendList"+tx+
				" value="+tx+"裁決寄存送達清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funSendArriveBat();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_SendDetailList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_SendDetailList"+tx+
				" value="+tx+"移送案件明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserSend_Stop();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayDetailList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayDetailList"+tx+
				" value="+tx+"繳費明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserPay_Stop();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReport="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReport"+tx+
				" value="+tx+"收繳費統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funUnitListExecel();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_NoPayDetailList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NoPayDetailList"+tx+
				" value="+tx+"未繳費明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserNoPay_Stop();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_NoPayReport="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NoPayReport"+tx+
				" value="+tx+"未繳費統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funNoPayUnitListExecel();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_YearSaveList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_YearSaveList"+tx+
				" value="+tx+"年度保留清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasser_Reserve();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_UrgeList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_UrgeList"+tx+
				" value="+tx+"催告清冊"+tx+" class="+tx+"btn3"+tx+" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funSendExecel()"+tx+">"+
					"&nbsp;&nbsp;";

var btn_UrgeCloseList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_UrgeCloseList"+tx+
				" value="+tx+"催告已到案清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funSendCloseExecel();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_UrgeNotCloseList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_UrgeNotCloseList"+tx+
				" value="+tx+"催告未到案清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funUcloseExecel();"+tx+">"+
					"&nbsp;&nbsp;";


var btn_NotArrivedList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NotArrivedList"+tx+
				" value="+tx+"未送達清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPassersEndArrivedList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_NoPayExeList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NoPayExeList"+tx+
				" value="+tx+"未繳納待執行清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funNoSendList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_SendnotCreditList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_SendnotCreditList"+tx+
				" value="+tx+"已移送未執行債權清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserSendnotCredit();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_SendtwoList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_SendtwoList"+tx+
				" value="+tx+"債權憑證準備再移送清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserSendtwoList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_InventoryList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_InventoryList"+tx+
				" value="+tx+"交付保管品核對清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserInventoryList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_NoEffectsList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NoEffectsList"+tx+
				" value="+tx+"無個人財產清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funNoEffectsList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_CreditorList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_CreditorList"+tx+
				" value="+tx+"債權憑證清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funCreditorList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_CreditorStstusList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_CreditorStstusList"+tx+
				" value="+tx+"債權執行狀態清冊"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funCreditorStstusList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_SendCreditorList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_SendCreditorList"+tx+
				" value="+tx+"債權明細統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserSendCreditorList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_SendCreditorTwoList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_SendCreditorTwoList"+tx+
				" value="+tx+"債權再移送明細統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserSendCreditorTwoList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayDetail_YiLanList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayDetail_YiLanList"+tx+
				" value="+tx+"行政罰鍰收繳情形明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserPayDetail_YiLanList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_NoPayDetail_YiLanList="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_NoPayDetail_YiLanList"+tx+
				" value="+tx+"應收未收收繳情形明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserNoPayDetail_YiLanList();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayCityReport="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayCityReport"+tx+
				" value="+tx+"(署)交通管理事件統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReport();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReportMonth="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReportMonth"+tx+
				" value="+tx+"交通罰緩收入憑證月報表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReportMonth();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReportMonthUnit="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReportMonthUnit"+tx+
				" value="+tx+"(分局)交通罰緩收入憑證月報表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReportMonth_Unit();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReportDayUnit="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReportDayUnit"+tx+
				" value="+tx+"(分局)交通罰緩收據明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReportDay_Unit();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReport2="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReport2"+tx+
				" value="+tx+"各單位清理統計表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReport2();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReport3="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReport3"+tx+
				" value="+tx+"各單位清理進度表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPayReport3();"+tx+">"+
					"&nbsp;&nbsp;";

var btn_PayReport4="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+
				"<input type="+tx+"button"+tx+" name="+tx+"btn_PayReport4"+tx+
				" value="+tx+"交通違規罰鍰繳款明細表"+tx+" class="+tx+"btn3"+tx+
				" style="+tx+"width:80%;height:30px;font-size:16px;line-height:20px;"+tx+
				" onclick="+tx+"myForm.chk_button_event.value=1;funPasserPayCloseDetail();"+tx+">"+
					"&nbsp;&nbsp;";




function ExePross(selt_sn){

	if(eval("exe_menu_"+selt_sn).style.visibility=='hidden'){

		if(myForm.chk_button_Other.value!=""){

			eval(myForm.chk_button_Other.value).style.visibility='hidden';
			myForm.chk_button_Other.value="";
		}

		eval("exe_menu_"+selt_sn).style.left=event.clientX+20+'px';
		eval("exe_menu_"+selt_sn).style.top=event.clientY-10+'px';
		eval("exe_menu_"+selt_sn).style.visibility='visible';

		myForm.chk_button_Other.value="exe_menu_"+selt_sn;

	}else{
		eval("exe_menu_"+selt_sn).style.visibility='hidden';
		myForm.chk_button_Other.value="";
	}
}

function AddReportMenu(){
	if(prt_menu.style.visibility=='hidden'){
		prt_menu.style.position='static';
		prt_menu.style.visibility='visible';
		prt_menu.style.width='300px';
	}else{
		prt_menu.style.width='0%';
		prt_menu.style.position='absolute';
		prt_menu.style.visibility='hidden';
	}
}

function AddQuery(){
	if(menu.style.visibility=='hidden'){
		menu.style.left=document.body.clientWidth/2-300+'px';
		//menu.style.top=document.body.scrollTop+event.clientY-5;
		menu.style.top=document.body.scrollTop+10;
		menu.style.visibility='visible';
	}else{
		menu.style.visibility='hidden';
	}
}

function DelQuery(){
	var tmpObjName=myForm.Selt_SpanName.value.split(',');
	var tmpObj=myForm.Selt_Span.value.split(',');

	for(j=1;j<=tmpObjName.length-2;j++){

		var tmpQryName=tmpObjName[j].split('@');
		var tmpQry=tmpObj[j].split('@');

		if(eval(tmpQryName[0]).style.backgroundColor=='#cccccc'){

			for(i=0;i<=myForm.seleObj.length-1;i++){

				if(myForm.seleObj[i].checked){
					if(myForm.seleObj[i].value==tmpQry[1]){

						eval(tmpQryName[0]).style.backgroundColor='#ffffff';
						eval(tmpQry[0]).style.backgroundColor='#ffffff';

						myForm.seleObj[i].click();
						break;
					}
				}
			}
			break;
		}
	}


	var tmpObj=myForm.Selt_Rpt.value.split(',');

	for(j=1;j<=tmpObj.length-2;j++){

		var tmpQry=tmpObj[j].split('@');

		if(eval(tmpQry[0]).style.backgroundColor=='#cccccc'){

			for(i=0;i<=myForm.RptSelet.length-1;i++){

				if(myForm.RptSelet[i].checked){
					if(myForm.RptSelet[i].value==tmpQry[1]){

						eval(tmpQry[0]).style.backgroundColor='#ffffff';

						myForm.RptSelet[i].click();
						break;
					}
				}
			}
			break;
		}
	}
}

function Check_Rpt(obj,QryObj){

	if(obj.checked){

		var tmpObj=myForm.chk_Rpt.value.split(',');

		for(j=1;j<=tmpObj.length-2;j++){

			if(eval(tmpObj[j]).innerHTML==""){

				myForm.Selt_Rpt.value+=tmpObj[j]+'@'+QryObj+',';

				eval(tmpObj[j]).innerHTML=eval(QryObj);

				break;
			}

		}
		
	}else{

		var tmpObj=myForm.Selt_Rpt.value.split(',');

		for(j=1;j<=tmpObj.length-2;j++){

			if(tmpObj[j].search('@'+QryObj)>=0){

				var tmpQry=tmpObj[j].split('@');

				myForm.Selt_Rpt.value=myForm.Selt_Rpt.value.replace(','+tmpQry[0]+'@'+QryObj+',', ",");

				eval(tmpQry[0]).innerHTML="";

				break;
			}

		}
	}
	
	//alert(myForm.Selt_Rpt.value);
		/*alert(myForm.chk_Span.value);
		alert(myForm.Selt_SpanName.value);
		alert(myForm.Selt_Span.value);*/

}

function Check_Selt(obj,QryName,QryObj){

	if(obj.checked){

		var tmpObjName=myForm.chk_SpanName.value.split(',');
		var tmpObj=myForm.chk_Span.value.split(',');

		for(j=1;j<=tmpObjName.length-2;j++){

			if(eval(tmpObjName[j]).innerHTML==""){

				myForm.Selt_SpanName.value+=tmpObjName[j]+'@'+QryName+',';
				myForm.Selt_Span.value+=tmpObj[j]+'@'+QryObj+',';

				eval(tmpObjName[j]).innerHTML=eval(QryName);
				eval(tmpObj[j]).innerHTML=eval(QryObj);

				break;
			}

		}
		
	}else{

		var tmpObjName=myForm.Selt_SpanName.value.split(',');
		var tmpObj=myForm.Selt_Span.value.split(',');

		for(j=1;j<=tmpObjName.length-2;j++){

			if(tmpObjName[j].search('@'+QryName)>=0){

				var tmpQryName=tmpObjName[j].split('@');
				var tmpQry=tmpObj[j].split('@');

				myForm.Selt_SpanName.value=myForm.Selt_SpanName.value.replace(','+tmpQryName[0]+'@'+QryName+',', ",");
				myForm.Selt_Span.value=myForm.Selt_Span.value.replace(','+tmpQry[0]+'@'+QryObj+',', ",");

				
				eval(tmpQryName[0]).innerHTML="";
				eval(tmpQry[0]).innerHTML="";

				break;
			}

		}
	}
	/*
	alert(myForm.chk_SpanName.value);
		alert(myForm.chk_Span.value);
		alert(myForm.Selt_SpanName.value);
		alert(myForm.Selt_Span.value);
	*/
}


function ChangeRpt(obj){

	var tmpObj=myForm.chk_Rpt.value.split(',');

	var strName="";
	
	if(myForm.chk_button_event.value==1){
		myForm.chk_button_event.value="";
		return false;
	}

	for(j=1;j<=tmpObj.length-2;j++){

		if(tmpObj[j]==obj){

			strName=tmpObj[j];

			break;
		}

	}

	if(eval(obj).style.backgroundColor=='#ffffff'){
		var chkValue=0;

		for(j=1;j<=tmpObj.length-2;j++){
			
			if(eval(tmpObj[j]).style.backgroundColor=='#cccccc'){

				var objTempID=eval(strName).innerHTML;

				//alert(myForm.Selt_SpanName.value);
				//alert(myForm.Selt_Span.value);

				eval(strName).innerHTML=eval(tmpObj[j]).innerHTML;

				eval(tmpObj[j]).innerHTML=objTempID;

				eval(tmpObj[j]).style.backgroundColor='#ffffff';

				myForm.Selt_Rpt.value=myForm.Selt_Rpt.value.replace(','+strName+'@', ",TTT@");

				myForm.Selt_Rpt.value=myForm.Selt_Rpt.value.replace(','+tmpObj[j]+'@', ','+strName+'@');

				myForm.Selt_Rpt.value=myForm.Selt_Rpt.value.replace(',TTT@', ','+tmpObj[j]+'@');

				//alert(myForm.Selt_SpanName.value);
				//alert(myForm.Selt_Span.value);

				chkValue=1;

				break;
			}

		}
		
		if(chkValue==0){
			eval(strName).style.backgroundColor='#cccccc';
		}
	}else{

		eval(strName).style.backgroundColor='#ffffff';
	}

	//alert(myForm.Selt_Rpt.value);
}


function ChangeLocation(obj){
	var tmpObjName=myForm.chk_SpanName.value.split(',');
	var tmpObj=myForm.chk_Span.value.split(',');

	var objID="";
	var OBjName="";

	for(j=1;j<=tmpObjName.length-2;j++){

		if(tmpObjName[j]==obj){

			strID=tmpObjName[j];
			strName=tmpObj[j];

			break;
		}

	}


	if(eval(obj).style.backgroundColor=='#ffffff'){

		var chkValue=0;

		for(j=1;j<=tmpObjName.length-2;j++){
			
			if(eval(tmpObjName[j]).style.backgroundColor=='#cccccc'){
				var objTempName=eval(strID).innerHTML;
				var objTempID=eval(strName).innerHTML;

				//alert(myForm.Selt_SpanName.value);
				//alert(myForm.Selt_Span.value);

				eval(strID).innerHTML=eval(tmpObjName[j]).innerHTML;
				eval(strName).innerHTML=eval(tmpObj[j]).innerHTML;

				eval(tmpObjName[j]).innerHTML=objTempName;
				eval(tmpObj[j]).innerHTML=objTempID;

				eval(tmpObjName[j]).style.backgroundColor='#ffffff';
				eval(tmpObj[j]).style.backgroundColor='#ffffff';

				myForm.Selt_SpanName.value=myForm.Selt_SpanName.value.replace(','+strID+'@', ',TTT@');
				myForm.Selt_Span.value=myForm.Selt_Span.value.replace(','+strName+'@', ",TTT@");

				myForm.Selt_SpanName.value=myForm.Selt_SpanName.value.replace(','+tmpObjName[j]+'@', ','+strID+'@');
				myForm.Selt_Span.value=myForm.Selt_Span.value.replace(','+tmpObj[j]+'@', ','+strName+'@');

				myForm.Selt_SpanName.value=myForm.Selt_SpanName.value.replace(',TTT@', ','+tmpObjName[j]+'@');
				myForm.Selt_Span.value=myForm.Selt_Span.value.replace(',TTT@', ','+tmpObj[j]+'@');

				//alert(myForm.Selt_SpanName.value);
				//alert(myForm.Selt_Span.value);

				chkValue=1;

				break;
			}

		}
		
		if(chkValue==0){
			eval(strID).style.backgroundColor='#cccccc';
			eval(strName).style.backgroundColor='#cccccc';
		}
	}else{

		eval(strID).style.backgroundColor='#ffffff';
		eval(strName).style.backgroundColor='#ffffff';
	}
}


function funSearchCname(inObj,CrObj){
	if(myForm.Sys_Driver.value!=''){
		runServerScript("Search_ChName.asp?inObj="+inObj+"&CrObj="+CrObj+"&chName="+myForm.Sys_Driver.value);
	}
}

function funCrtVale(inObj,CrObj,objValue){

	eval(CrObj).innerHTML="";

	if(objValue!=''){
		eval("document.all."+inObj).value=objValue;
	}
}



function funSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_Display.value='show';
	myForm.Sys_SendBillSN.value='';
	myForm.DB_KindSelt.value='';
	myForm.submit();
}
function funLoadSend(){
	if(document.getElementsByName("chkSend").length>0){
		for(i=0;i<=ck_Sn.length-1;i++){
			for(j=0;j<=myForm.chkSend.length-1;j++){
				if(ck_Sn[i]==myForm.chkSend[j].value){
					myForm.chkSend[j].checked=true;
				}
			}
		}
	}

	<%=js_obj%>
}
function funChkSend(){
	var tempSend='';
	var chked=false;
	for(j=0;j<=myForm.chkSend.length-2;j++){
		if(myForm.chkSend[j].checked){
			chked=false
			for(i=0;i<=ck_Sn.length-1;i++){
				if(myForm.chkSend[j].value==ck_Sn[i]){
					chked=true;
				}
			}
			if(!chked){
				if(strBillSN!=''){
					strBillSN=strBillSN+',';
				}
				strBillSN=strBillSN+myForm.chkSend[j].value;
			}
		}
	}

	myForm.Sys_SendBillSN.value=strBillSN;
	ck_Sn=strBillSN.split(',');
	for(i=0;i<=ck_Sn.length-1;i++){
		chked=false
		for(j=0;j<=myForm.chkSend.length-1;j++){
			if(!myForm.chkSend[j].checked){
				if(myForm.chkSend[j].value==ck_Sn[i]){
					chked=true;
				}
			}
		}
		
		if(!chked){
			if(tempSend!=''){
				tempSend=tempSend+',';
			}
			tempSend=tempSend+ck_Sn[i];
		}
	}
	myForm.Sys_SendBillSN.value=tempSend;
	strBillSN=myForm.Sys_SendBillSN.value;
	ck_Sn=strBillSN.split(',');
}

function funPrintPeopleList_Stop(){
	UrlStr="../Query/PrintPeopleDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPasserPay_Stop(){
	UrlStr="PasePayDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funPasserSend_Stop(){
	UrlStr="PasserSendDetail_Execel.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPasserNoPay_Stop(){
	UrlStr="PaseNoPayDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReportMonth(){
	var error=0;
	if(myForm.PayDate1.value!=""){
		if(!dateCheck(myForm.PayDate1.value)){
			error=1;
			alert("付費日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入付費日!!");
	}
	if (error==0){
		if(myForm.PayDate2.value!=""){
			if(!dateCheck(myForm.PayDate2.value)){
				error=1;
				alert("付費日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入付費日!!");
		}
	}
	if (error==0){
		myForm.action="ReportMonthPay.asp";
		myForm.target="PayReportMonth";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
		
}

function funPayReportMonth_Unit(){
	var error=0;
	if(myForm.PayDate1.value!=""){
		if(!dateCheck(myForm.PayDate1.value)){
			error=1;
			alert("付費日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入付費日!!");
	}
	if (error==0){
		if(myForm.PayDate2.value!=""){
			if(!dateCheck(myForm.PayDate2.value)){
				error=1;
				alert("付費日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入付費日!!");
		}
	}
	if (error==0){
		myForm.action="ReportMonthPay_Unit.asp";
		myForm.target="PayReportMonth";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
		
}

function funPayReportDay_Unit(){
	var error=0;
	if(myForm.PayDate1.value!=""){
		if(!dateCheck(myForm.PayDate1.value)){
			error=1;
			alert("付費日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入付費日!!");
	}
	if (error==0){
		if(myForm.PayDate2.value!=""){
			if(!dateCheck(myForm.PayDate2.value)){
				error=1;
				alert("付費日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入付費日!!");
		}
	}

	if (error==0){
		myForm.action="ReportMonthDay_Unit.asp";
		myForm.target="ReportMonthDay_Unit";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
		
}

function funPayReport(){

	myForm.action="../Report/Report0110.asp";
	myForm.target="funPayReport";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReport2(){

	myForm.action="../Report/Report0113.asp";
	myForm.target="funPayReport2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReport3(){

	myForm.action="../Report/Report0114.asp";
	myForm.target="funPayReport3";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPasserPayCloseDetail(){
	
	myForm.action="PasserPayCloseDetail.asp";
	myForm.target="PasserPayCloseDetail";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
}

function funSltUrgeDate(){
	var error=0;
	/*if(myForm.Sys_sltUrgeDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltUrgeDate1.value)){
			error=1;
			alert("催告日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入催告日!!");
	}
	if (error==0){
		if(myForm.Sys_sltUrgeDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltUrgeDate2.value)){
				error=1;
				alert("催告日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入催告日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltUrgeDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funSltJudeDate(){
	var error=0;
	/*if(myForm.Sys_sltJudeDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltJudeDate1.value)){
			error=1;
			alert("裁決日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入裁決日!!");
	}
	if (error==0){
		if(myForm.Sys_sltJudeDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltJudeDate2.value)){
				error=1;
				alert("裁決日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入裁決日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltJudeDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funSltSendDate(){
	var error=0;
	/*if(myForm.Sys_sltSendDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltSendDate1.value)){
			error=1;
			alert("移送日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入移送日!!");
	}
	if (error==0){
		if(myForm.Sys_sltSendDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltSendDate2.value)){
				error=1;
				alert("移送日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入移送日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltSendDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funUrgeDateSelt(){
	var error=0;
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="UrgeDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funJudeDateSelt(){
	var error=0;
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="JudeDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funSendDateSelt(){
	var error=0;
	myForm.hd_BillSN.value='';
	myForm.hd_BillNo.value='';
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="SendDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funPrintStyle(){

		UrlStr="PasserSendStyle.asp";
		newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
		myForm.action="PasserSendStyle.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funsubmit(){
	PasserWin.close();
	if(myForm.printStyle.value=='0'){
		UrlStr="BillPrints.asp";
	}else{
		UrlStr="BillPrints_a4.asp";
	}
	newWin(UrlStr,"JudeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="JudeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}
function funchgprint(){
	PasserWin.DP();
}

function funUrgeList(){
	UrlStr="BillPrints_legal.asp";
	newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="UrgeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}

function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

function UnitMan(UnitName,UnitMem,tmpMemID){
	if(document.all[UnitName].value!=''){
		runServerScript("UnitAdd.asp?UnitID="+document.all[UnitName].value+"&UnitMem="+UnitMem+"&MemberID="+tmpMemID);
	}else{
		document.all[UnitMem].options[1]=new Option('請選取','');
		document.all[UnitMem].length=1;
	}
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			
			myForm.hd_BillSN.value='';
			myForm.hd_BillNo.value='';
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){

			myForm.hd_BillSN.value='';
			myForm.hd_BillNo.value='';
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	PasserWin=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	PasserWin.focus();
	return win;
}
function funchgExecel(){
	if(myForm.DB_Display.value=="show"){
		//newWin("","PasserBaseQry_Execel",980,550,0,0,"yes","yes","yes","no");
		UrlStr="PasserBaseQry_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="PasserBaseQry_Execel";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseJudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseSendList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="JudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendCloseExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseSendCloseList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funPasserMailMoney(){
	if(myForm.DB_Display.value=="show"){
		newWin("PasserMailMoneyList_Select.asp","PasserMailMoneyList",400,200,50,10,"yes","yes","yes","no");
	}else{
		alert("請先進行查詢");
	}
}
function funPasserReportList(){
	UrlStr="PasserReportList.asp";
	myForm.action=UrlStr;
	myForm.target="PasserReportList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funUcloseExecel(){
	if(myForm.DB_Display.value=="show"){
		myForm.action="PasserBaseSendUCloseList_Execel.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funPasser_Reserve(){
	if(myForm.DB_Display.value=="show"){
		if(Sys_City=='台南市'){
			myForm.action="PasserBaseReserveList_TaiNaNCity.asp";
		}else{
			myForm.action="PasserBaseReserveList.asp";
		}

		//myForm.action="PasserBaseReserveList.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funUrgeJudeExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseUrgeJudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funUnitListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseUnitList_Execel.asp"
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funNoPayUnitListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseNotPayUnitList_Execel.asp"
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funCountryBat(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserCountryList.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funCreditorStstusList(){
    if(myForm.DB_Display.value=="show"){
		
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserCreditorStstusList.asp"
		myForm.action=UrlStr;
		myForm.target="CreditorStstusList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
    }else{
        alert("請先進行查詢");
    }
}
function funPasserSendCreditorList(){
		
	//UrlStr="PasserArriveList.asp";
	UrlStr="PasserSendCreditorList.asp"
	myForm.action=UrlStr;
	myForm.target="PasserSendCreditorList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funPasserSendCreditorTwoList(){
		
	//UrlStr="PasserArriveList.asp";
	UrlStr="PasserSendCreditorTwoList.asp"
	myForm.action=UrlStr;
	myForm.target="PasserSendCreditorTwoList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funCreditorList(){
    if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserCreditorNotPayList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserCreditorNotPayList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
    }else{
        alert("請先進行查詢");
    }
}

function funPassersEndArrivedList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="NotArrivedList.asp"
		myForm.action=UrlStr;
		myForm.target="NotArrivedList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funNoSendList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="NoSendList.asp"
		myForm.action=UrlStr;
		myForm.target="NoSendList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserSendnotCredit(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserSendnotCredit.asp"
		myForm.action=UrlStr;
		myForm.target="PasserSendnotCredit";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserSendtwoList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserSendtwoList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserSendtwoList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserInventoryList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserInventoryList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserInventoryList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserPayDetail_YiLanList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserPayDetail_YiLanList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserPayDetail_YiLanList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserNoPayDetail_YiLanList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserNoPayDetail_YalinList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserNoPayDetail_YalinList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funNoEffectsList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserNoEffectsList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funGovArriveBat(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="funGovSendList_Exce.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendArriveBat(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="funSendSendList_Exce.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserJudeBat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funJudeBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserJudeBat_chromat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funSendBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserSendDetailBat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funSendBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserSendDetailBat_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funSendBatTwo_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;
		
		myForm.action="PasserSendDetailBatTwo_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funUrgeBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserUrgeBat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funUrgeBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserUrgeBat_chromat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funJudeList(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		myForm.action="PasserJudeBatList.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funJudeList_chromat(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		<%if sys_City="台中市" then
			Response.Write "myForm.action=""PasserJudeBatList_chromat.asp"";"
		elseif sys_City="台中縣" then
			Response.Write "myForm.action=""PasserJudeBatListTaiChung_chromat.asp"";"
		end if%>
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendList(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		myForm.action="PasserSendBatList.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendList_chromat(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		<%if sys_City="台中市" then
			Response.Write "myForm.action=""PasserSendBatList_chromat.asp"";"
		elseif sys_City="台中縣" then
			Response.Write "myForm.action=""PasserSendBatListTaiChung_chromat.asp"";"
		end if%>
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendListTwo_chromat(SendType){
		myForm.Sys_PasserNotify.value=SendType;
		myForm.action="PasserSendBatListTwo_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funRecallData(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserRecallData.asp";
		myForm.action=UrlStr;
		myForm.target="ReCall";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funUpdAddress(){
		UrlStr="PasserUpdAddress.asp";
		myForm.action=UrlStr;
		myForm.target="PasserUpdAddress";
		myForm.submit();
		myForm.action="";
		myForm.target="";

}
</script>
<%
conn.close
set conn=nothing
%>