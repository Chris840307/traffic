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
	strSQL="update PasserBase set forfeit2=(select nvl(max(level2),0) from law where version=2 and itemid=passerbase.rule2) where rule2 is not null and nvl(forfeit2,0)=0 and (select count(1) cnt from passerJude where billsn=passerbase.sn)>0 and recordDate between to_date(TO_CHAR(SYSDATE-750, 'YYYY/MM/DD')||' 00:00:00','YYYY/MM/DD/HH24/MI/SS') and to_date(TO_CHAR(SYSDATE, 'YYYY/MM/DD')||' 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

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
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
   font-weight:900;
}

.err01 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
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
'增加時要改報表 Report0144.asp
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then 

	showCreditor=true 
end If

If isempty(request("DB_Selt")) Then

	'strSQL="delete passersenddetail where exists(select 'Y' from (select sendnumber,senddate from (select nvl(sendnumber,'1') sendnumber,nvl(senddate,to_date('1999/01/01','YYYY/MM/DD')) senddate,count(1) cnt from passersenddetail group by nvl(sendnumber,'1'),nvl(senddate,to_date('1999/01/01','YYYY/MM/DD')) ) tba where cnt >1 ) tab where sendnumber=nvl(passersenddetail.sendnumber,'1') and senddate=nvl(passersenddetail.senddate,to_date('1999/01/01','YYYY/MM/DD'))) and Not Exists(select 'N' from PASSERCREDITOR where senddetailsn=passersenddetail.SN)"

	'conn.execute(strSQL)

	'strSQL="delete PasserCreditor where exists(select 'Y' from passerbase where sn=PasserCreditor.billsn) and not exists(select 'Y' from passersenddetail where sn=PasserCreditor.senddetailsn)"

	'conn.execute(strSQL)

	strSQL="select * from PassersEndArrived where rownum=1"
	set rs=conn.execute(strSQL)
	If Not rs.eof Then
		For i=0 to rs.Fields.count-1
			If trim(rs.Fields.item(i).Name)="MAILDATE" Then Exit For
		Next
		If i>rs.Fields.count-1 Then
			strSQL="Alter Table PassersEndArrived ADD (MAILDATE DATE)"
			conn.execute(strSQL)
		End if
	End if

End if 

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

'==========================================================================================


'==========================================================================================

If DB_KindSelt = "Cancel" Then
	theBatchTime=""
	strSN="select PASSERDCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"C"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing	

	DCISN=""
	strDciSN="select passerDCILOG_SEQ.nextval as SN from Dual"
		set rsSN=conn.execute(strDciSN)
		if not rsSN.eof then
			DCISN=trim(rsSN("SN"))
		end if
		rsSN.close
	set rsSN=nothing
	
	strInsCaseIn="insert into PASSERDCILOG(" & _
				"SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" & _
				",RecordMemberID,ExchangeDate,ReturnMarkType,ExchangeTypeID,BatchNumber,DciUnitID)"&_
				"values(" & DCISN & ","&Request("hd_CanCelSN")&","&_
				"(select billno from passerbase where sn="&Request("hd_CanCelSN")&")"&_
				",(select billtypeid from passerbase where sn="&Request("hd_CanCelSN")&")"&_
				",(select CarNo from passerbase where sn="&Request("hd_CanCelSN")&")"&_
				",(select billunitid from passerbase where sn="&Request("hd_CanCelSN")&")"&_
				",sysdate,"&Session("User_ID")&",sysdate,8,'C','"&theBatchTime&"'" &_
				",(" &_
					"select DciUnitID from UnitInfo ut where UnitID=(" &_
						"select unittypeid from UnitInfo uta where unitid=(select billunitid from passerbase where sn="&Request("hd_CanCelSN")&")" &_
					")" &_
				")" &_
			")" 

	conn.execute strInsCaseIn

	strInsCaseIn="insert into PasserBaseDciReturn(" & _
		"DciLogSN,BillSN,BillNO,CarNo,ExchangeTypeID)"&_
		"values(" & DCISN & ","&Request("hd_CanCelSN")&","&_
		"(select billno from passerbase where sn="&Request("hd_CanCelSN")&"),"&_
		"(select CarNo from passerbase where sn="&Request("hd_CanCelSN")&"),'C')" 

	conn.execute strInsCaseIn



	sqlpasserbase="update PasserBase set BillStatus=2,DCILOGSN=(select max(sn) from PASSERDCILOG where billsn=passerbase.sn and exchangetypeid='W') where sn="&Request("hd_CanCelSN")

	conn.execute sqlpasserbase

End if 

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

		strwhere=strwhere&" and Exists(select 'Y' from PasserSendDetail psd01 where billsn=a.sn and (select max(SendDate) from PasserSendDetail psd02 where billsn=a.sn) between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')) and (select count(1) cnt from PasserSendDetail dt where billsn=a.sn)>1"
	end If  
	
'	if request("Sys_CreditorType")<>"" then
'		
'		strwhere=strwhere&" and Not Exists(select 'Y' from PasserCreditor where billsn=a.sn)"
'	end If

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

	if trim(request("sys_CreditorTypeID"))<>"" Or (trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"") Or (trim(request("max_PetitionDate1"))<>"" And trim(request("max_PetitionDate2"))<>"") Then
		strPasserCreditorAdd=""
		If trim(request("sys_CreditorTypeID")) <> "-1" Then
			If trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"" Then
				PetitionDate1=gOutDT(request("Sys_PetitionDate1"))&" 0:0:0"
				PetitionDate2=gOutDT(request("Sys_PetitionDate2"))&" 23:59:59"

				If strPasserCreditorAdd="" Then

					strPasserCreditorAdd=" and (select min(PetitionDate) from PasserCreditor where Billsn=a.sn) between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				
				End If 
			End If 

			If trim(request("max_PetitionDate1"))<>"" And trim(request("max_PetitionDate2"))<>"" Then
				PetitionDate1=gOutDT(request("max_PetitionDate1"))&" 0:0:0"
				PetitionDate2=gOutDT(request("max_PetitionDate2"))&" 23:59:59"

				strPasserCreditorAdd=strPasserCreditorAdd&" and (select max(PetitionDate) from PasserCreditor pc1 where Billsn=a.sn) between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS') and (select count(1) from PasserCreditor pc2 where Billsn=a.sn)>1"

			End If 
			
			If trim(request("sys_CreditorTypeID"))<>"" Then
					
				strPasserCreditorAdd=strPasserCreditorAdd&" and CreditorTypeID in('"&trim(request("sys_CreditorTypeID"))&"')"
			End If 
		end if

		If trim(request("sys_CreditorTypeID")) = "-1" Then
			strwhere=strwhere&" and not Exists(select 'N' from PasserCreditor where billsn=a.sn)"
		else
			strwhere=strwhere&" and Exists(select 'Y' from PasserCreditor where billsn=a.sn "&strPasserCreditorAdd&")"
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


	if request("Sys_CarNo")<>"" then
		strwhere=strwhere&" and a.CarNo='"&request("Sys_CarNo")&"'"
	end If 
 
	'違規人身分証號
	if request("Sys_DriverID")<>"" then
		strwhere=strwhere&" and a.DriverID='"&Ucase(request("Sys_DriverID"))&"'"
	end if

	if request("Sys_BillMemID")<>"" then
		strwhere=strwhere&" and a.BillMemID1 in(select MemberID from MemberData where LoginID='"&trim(request("Sys_BillMemID"))&"')"
	end if

	if request("Sys_SendMailStation")<>"" then
		strwhere=strwhere&" and (select count(1) cnt from PassersEndArrived where SendMailStation like '"&trim(Request("Sys_SendMailStation"))&"%' and PasserSN=a.sn)>0"
	end if

	if request("Sys_BILLSTATUS")<>"" then
		if trim(request("Sys_BILLSTATUS"))="9" then
			strwhere=strwhere&" and a.BILLSTATUS=9"
		
		elseif trim(request("Sys_BILLSTATUS"))="1" then
			strwhere=strwhere&" and a.BILLSTATUS=9 and exists(select 'Y' from PasserPay where billsn=a.sn and PayAmount>0)"

		elseif trim(request("Sys_BILLSTATUS"))="2" then
			strwhere=strwhere&" and a.BILLSTATUS=9 and (select nvl(sum(PayAmount),0) from PasserPay where billsn=a.sn)=0"

		elseif trim(request("Sys_BILLSTATUS"))="3" then
			strwhere=strwhere&" and (select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0"

		elseif trim(request("Sys_BILLSTATUS"))="4" then
			strwhere=strwhere&"and exists(select 'Y' from PasserPay where billsn=a.sn)"

		elseif trim(request("Sys_BILLSTATUS"))="5" then
			strwhere=strwhere&"and (a.BILLSTATUS<>9 or (select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0)"

		else
			strwhere=strwhere&" and a.BILLSTATUS<>9 and (select count(1) from passerpay where billsn=a.sn)=0"
		end if
	end if

	if request("Sys_Rule")<>"" then
		strwhere=strwhere&" and (Rule1 like '"&trim(request("Sys_Rule"))&"%' or Rule2 like '"&trim(request("Sys_Rule"))&"%' or Rule3 like '"&trim(request("Sys_Rule"))&"%' or Rule4 like '"&trim(request("Sys_Rule"))&"%')"
	end If 
	
	if trim(request("Sys_Fastener1"))<>"" then
		strwhere=strwhere&" and exists(select 'Y' from PasserConfiscate where ConfiscateID='"&trim(request("Sys_Fastener1"))&"' and billsn=a.sn)"
	end if 

	if trim(request("Sys_SendNumber"))<>"" then
		strwhere=strwhere&" and (exists(select 'Y' from PasserSend where SendNumber like '%"&trim(request("Sys_SendNumber"))&"%' and billsn=a.sn) or exists(select 'Y' from PasserSendDetail where SendNumber like '%"&trim(request("Sys_SendNumber"))&"%' and billsn=a.sn))"
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

	if request("Sys_ProjectID")<>"" then

		strwhere=strwhere&" and a.ProjectID in('"&request("Sys_ProjectID")&"')"
	end If 
	
	if request("Sys_ReserveYear1")<>"" and request("Sys_ReserveYear2")<>"" then

		strwhere=strwhere&" and a.ReserveYear between "&request("Sys_ReserveYear1")&" and "&request("Sys_ReserveYear2")
	end If 
		
	if request("Sys_checkdate1")<>"" and request("Sys_checkdate2")<>"" then

		strwhere=strwhere&" and Exists(select 'Y' from PasserProperty where billsn=a.sn and to_char(checkdate,'YYYY') between '"&(cdbl(request("Sys_checkdate1"))+1911)&"' and '"&(cdbl(request("Sys_checkdate2"))+1911)&"')"
	end If 
	
	if trim(request("Sys_DriverType"))="1" then

		strwhere=strwhere&" and not (length(driverid)=10 and substr(driverid,1,1) between 'A' and 'Z' and substr(driverid,2,1) between '1' and '2' and substr(driverid,3,1) between '0' and '9')"

	elseif trim(request("Sys_DriverType"))="2" then

		strwhere=strwhere&" and (length(driverid)=10 and substr(driverid,1,1) between 'A' and 'Z' and substr(driverid,2,1) between '1' and '2' and substr(driverid,3,1) between '0' and '9')"
	end If 
	
	if request("Sys_batchnumber")<>"" then

		strwhere=strwhere&" and (select count(1) from passerdcilog where batchnumber in('"&UCase(trim(request("Sys_batchnumber")))&"') and billsn=a.sn)>0"
	end If 

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
		strwhere=strwhere&" and not Exists(select 'N' from PasserJude where BillSN=a.SN) and ((select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0 or (select count(1) from PasserPay where billsn=a.sn)=0)"
	end if

	if DB_KindSelt="JudeDateSelt" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where BillSN=a.SN) and not Exists(select 'N' from PasserUrge where BillSN=a.SN) and not Exists(select 'N' from PasserSend where BillSN=a.SN) and ((select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0 or (select count(1) from PasserPay where billsn=a.sn)=0)"
	end if

	if DB_KindSelt="SendDateSelt" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where BillSN=a.SN) and not Exists(select 'N' from PasserSend where BillSN=a.SN) and ((select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0 or (select count(1) from PasserPay where billsn=a.sn)=0)"
	end If 	

	if DB_KindSelt="PasserProperty" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserCreditor where billsn=a.SN) and ((select count(1) from PasserPay where billsn=a.sn and CaseCloseDate is null)>0 or (select count(1) from PasserPay where billsn=a.sn)=0)"
		strwhere=strwhere&" and not Exists(select 'N' from PasserProperty where billsn=a.sn and to_number(to_char(checkdate,'YYYY'))<to_number(to_char(sysdate,'YYYY')))"
	end If 

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
	if trim(strwhere)="" then
		DB_Selt="":DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end If 

Sys_SendBillSN=trim(request("Sys_SendBillSN"))

'if sys_City="台中市" and isempty(request("DB_Selt")) then
'
'	DB_Display="show"
'
'	If Session("UnitLevelID") > 1 then MemberStation=" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"
'
'	strwhere=strwhere&" and billstatus<>9 and recordstateid=0 and BillFillDate > to_date('"&(year(now)-6)&"/12/31','YYYY/MM/DD') and TRUNC(sysdate-DeallineDate)>=60 and (select count(1) cnt from PasserJude where billsn=a.SN)=0 "&MemberStation
'end If 

if DB_Display="show" then

	showFiled="":PasserSendFiled="(select min(SendDate) as SendDate from PasserSend where BillSn=a.sn and SendDate is not null)"
	If showCreditor Then
		showFiled=",(select max(PetitionDate) PetitionDate from PasserCreditor where billsn=a.SN) PetitionDate"
		PasserSendFiled="(select min(SendDate) as SendDate from PasserSendDetail where BillSn=a.sn and SendDate is not null)"
	End if 	

	If sys_City="彰化縣" then

		showFiled=showFiled&",(select checkdate from PasserProperty where billsn=a.sn and to_char(checkdate,'YYYY')=to_char(sysdate,'YYYY')) checkdate"
		
		showFiled=showFiled&",(select postNumber from PasserProperty where billsn=a.sn and to_char(checkdate,'YYYY')=to_char(sysdate,'YYYY')) postNumber"

	end If 

	show_postNumber=""
	
	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Carno,a.CARSIMPLEID,a.Driver,a.IllegalAddress,a.Rule1,a.Rule2," &_
	"a.RuleVer,a.FORFEIT1,a.FORFEIT2,a.FORFEIT3,a.FORFEIT4,a.BILLSTATUS," &_
	"a.BillMem1,a.DoubleCheckStatus,nvl(a.dcilogsn,0) dcilogsn," &_
	"(Select JudeDate from PasserJude where billsn=a.sn) JUDEDATE," &_
	PasserSendFiled&" SENDDATE," &_
	"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
	"(Select UrgeDate from PasserUrge where billsn=a.sn) URGEDATE," &_
	"(select Max(ArrivedDate) ArrivedDate from PassersEndArrived where PasserSN=a.sn) ArrivedDate," &_
	"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate"&showFiled &_
	" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	'set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere

	strSQL="select a.SN,a.BillNo,a.Carno,a.billtypeid,a.DriverAddress,a.DoubleCheckStatus,a.Driver,a.IllegalDate,a.Rule1,a.Rule2,"&PasserSendFiled&" SENDDATE "&showFiled&" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	
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

		chkOwnerAddr="":Sys_OwnerZipName=""

		If (not ifnull(rssn("Carno"))) and trim(rssn("billtypeid"))="2" Then
			If ifnull(rssn("DriverAddress")) Then
				
				carSQL="select Owner,ownerID,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rssn("SN"))&" and ExchangetypeID='A' and (select count(1) cnt from PasserDCILog where ExchangetypeID='A' and billsn="&trim(rssn("SN"))&" and dcireturnstatusid in(select dcireturn from dcireturnstatus where dcireturnstatus=1))>0"
	
				set rsfi=conn.execute(carSQL)

				if Not rsfi.eof then
					If Not ifnull(trim(rsfi("Owneraddress"))) Then
						chkOwnerAddr=trim(rsfi("Owneraddress"))&"(車)"

						strSQL="select ZipName from Zip where ZipID='"&trim(rsfi("OwnerZip"))&"'"
						set rszip=conn.execute(strSQL)
						if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
						rszip.close

						chkOwnerAddr=Sys_OwnerZipName&replace(replace(chkOwnerAddr,"臺","台"),Sys_OwnerZipName,"")

						strSQL="update passerbase set" & _
						" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
						",DriverID='"&trim(rsfi("ownerID"))&"',DriverAddress='"&chkOwnerAddr&"'"& _
						" where SN="&trim(rssn("SN"))

						conn.execute(strSQL)
					end if
				end If 
				rsfi.close

				If ifnull(chkOwnerAddr) Then					

					carSQL="select Owner,ownerID,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rssn("SN"))&" and ExchangetypeID='W' and (select count(1) cnt from PasserDCILog where ExchangetypeID='W' and billsn="&trim(rssn("SN"))&" and dcireturnstatusid in(select dcireturn from dcireturnstatus where dcireturnstatus=1))>0"
		
					set rsfi=conn.execute(carSQL)

					if Not rsfi.eof then
						If Not ifnull(trim(rsfi("Owneraddress"))) Then

							chkOwnerAddr=trim(rsfi("Owneraddress"))&"(車)"

							strSQL="select ZipName from Zip where ZipID='"&trim(rsfi("OwnerZip"))&"'"
							set rszip=conn.execute(strSQL)
							if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
							rszip.close

							chkOwnerAddr=Sys_OwnerZipName&replace(replace(chkOwnerAddr,"臺","台"),Sys_OwnerZipName,"")
							
							strSQL="update passerbase set" & _
							" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
							",DriverID='"&trim(rsfi("ownerID"))&"',DriverAddress='"&chkOwnerAddr&"'" & _
							" where SN="&trim(rssn("SN"))

							conn.execute(strSQL)
						end if
					end If 
					rsfi.close
				End if 
			end If 
		
		End if 

		rssn.movenext
	wend
	rssn.close

	strCnt="select count(1) as cnt from PasserBase a where RecordStateID=0 "&strwhere
	set Dbrs=conn.execute(strCnt)
	DBsum=0
	If Not Dbrs.eof Then
		DBsum=cdbl(Dbrs("cnt"))
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

	if sys_City="彰化縣" and isempty(request("DB_Selt")) then
		errmsg="":PasserCnt=0:PasserSendCnt=0:MeberStation=""

		If Session("UnitLevelID") > 1 then MemberStation=" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"

		strSQL="select (select UnitName from Unitinfo where UnitID=passerbase.MemberStation) MebUnitName,billno from passerbase where (select count(1) cnt from PasserSendDetail where billsn=passerbase.sn)=1 and exists(select 'Y' from PasserSendDetail where TRUNC(sysdate-SENDDATE)>1097 and billsn=passerbase.sn) and not Exists(select 'N' from PasserCreditor where PETITIONDATE is not null and BillSn=PasserBase.SN) and billstatus<>9 and TRUNC(sysdate-illegaldate)<3650 and recordstateid=0 "&MemberStation&" order by MebUnitName" 

		set rssn=conn.execute(strSQL)

		while Not rssn.eof


			errmsg=errmsg&rssn("MebUnitName")&",單號" & rssn("billno") & "移送逾三年未取得債權憑證\n"

			rssn.movenext
		wend
		
		rssn.close

		
		If not ifnull(errmsg) Then
			Response.write "<script>"
			Response.Write "alert('" & errmsg & "！');"
			Response.write "</script>"
		end if
	end If 

	if (sys_City="基隆市" or sys_City="屏東縣" or sys_City="台中市") and isempty(request("DB_Selt")) then
		
		errmsg="":PasserCnt=0:PasserSendCnt=0:PasserCanCel=0:MeberStation=""

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

		if sys_City="屏東縣" then

			strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and exists(select 'Y' from PasserJude where billsn=PasserBase.SN and TRUNC(sysdate-JudeDate)>3650) and Exists(select 'Y' from PasserCreditor where PETITIONDATE is not null and BillSn=PasserBase.SN)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"

			set rssn=conn.execute(strSQL)

			while Not rssn.eof
				PasserCanCel=cdbl(rssn("cnt"))

				if trim(PasserCanCel)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserCanCel & "筆已逾十年註銷案件。\n"

				rssn.movenext
			wend
			
			rssn.close
		end If  
		
		If not ifnull(errmsg) Then
			Response.write "<script>"
			Response.Write "alert('" & errmsg & "！');"
			Response.write "</script>"
		end if
		
	end If 

End if 

%>
<body onLoad="funLoadSend();">
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr height="30">
		<td height="30" bgcolor="#1BF5FF"><font size="4"><b>慢車行人道路障礙舉發單紀錄</b> </font><img src="space.gif" width="32" height="10"><a href="passerbase2.docx" target="_blank" ><font size="5"> 下載 裁罰系統使用說明.doc</font></a></img>
		<%if showCreditor then%>
			<B>
			<a href="PasserCreditor.doc" target="_blank" ><font size="5" color="red"> 下載 債權憑證系統使用說明.doc</font></a><br>
			<a href="PasserCreditorReport.doc" target="_blank" ><font size="5" color="red"> 下載 債權憑證報表使用說明.doc</font></a>
			</td>
			</b>
			
		<%end if%>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
						<tr><td>
						建檔序號
						</td><td>
						<input name="Sys_DoubleCheckStatus" class="btn1" type="text" value="<%=request("Sys_DoubleCheckStatus")%>" size="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						</td><td>
						舉發單位
						</td><td>
						<select name="Sys_BillUnitID" class="btn1"><%
							strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUnit=conn.execute(strSQL)
							If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
							rsUnit.close
							strUnitID=""
							if trim(Session("UnitLevelID"))="1" then
								strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
								strtmp=strtmp+"<option value="""">所有單位</option>"
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
								strtmp=strtmp+"<option value="""">所有單位</option>"
								'strtmp=strtmp+"<option value="""&strUnitID&""">管轄單位</option>"
							elseif trim(Session("UnitLevelID"))="3" then
								strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitTypeID,UnitName"
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
						<%'SelectUnitOption("Sys_BillUnitID","")%>
						
						<!--<img src="space.gif" width="3" height="10">
						舉發員警
						<%=SelectMemberOption("Sys_BillUnitID","Sys_BillMem")%>
						<img src="space.gif" width="3" height="10">-->
						<!--
						建檔人
						<select name="Sys_RecordMemberID" class="btn1">
							<option Value="">請選擇</option>
							<%
							strSQL="Select ChName ,MemberID from MemberData"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("MemberID")&""""
								if DB_Selt="Selt" then
									'if trim(rs1("MemberID"))=trim(request("Sys_RecordMemberID")) then response.write " selected"
								else
									if trim(rs1("MemberID"))=trim(session("User_ID")) then response.write " selected"
								end if	
								response.write ">"&rs1("ChName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						-->
						</td><td nowrap>
						舉發人背章號碼
						</td><td>
						<input name="Sys_BillMemID" class="btn1" type="text" value="<%=request("Sys_BillMemID")%>" size="7" maxlength="8">
						</td></tr><tr><td>
						<strong><font color="red">舉發單號</font></strong>
						</td><td>
						<input name="Sys_BillNo" maxlength="9" size="8" class="btn1" type="text" value="<%=Ucase(request("Sys_BillNo"))%>" size="8" maxlength="20">
						</td><td nowrap>
						應到案處
						</td><td>
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
						</td><td nowrap>
						繳費狀況
						</td><td>
						<select Name="Sys_BILLSTATUS" class="btn1">
							<option value="">全部</option>
							<option value="0"<%if trim(request("Sys_BILLSTATUS"))="0" then response.write " selected"%>>未繳費</option>
							<option value="5"<%if trim(request("Sys_BILLSTATUS"))="5" then response.write " selected"%>>未結案</option>
							<option value="4"<%if trim(request("Sys_BILLSTATUS"))="4" then response.write " selected"%>>已繳費(全部)</option>
							<option value="9"<%if trim(request("Sys_BILLSTATUS"))="9" then response.write " selected"%>>已繳費(已結案含免罰)</option>
							<option value="1"<%if trim(request("Sys_BILLSTATUS"))="1" then response.write " selected"%>>已繳費(已結案不含免罰)</option>
							<option value="2"<%if trim(request("Sys_BILLSTATUS"))="2" then response.write " selected"%>>已繳費(免罰)</option>
							<option value="3"<%if trim(request("Sys_BILLSTATUS"))="3" then response.write " selected"%>>已繳費(未結案)</option>
						</select>
						</td><td></tr><tr><td>
						違規人名
						</td><td>
						<input name="Sys_Driver" class="btn1" type="text" value="<%=request("Sys_Driver")%>" size="7" maxlength="8" onkeyup="funSearchCname('Sys_Driver','SearChName')" onMouseDown="funCrtVale('Sys_Driver','SearChName','');">
						<br>
						<div id="SearChName" style="position:absolute;">							
						</div>
						</td><td>
						身分證號
						</td><td>
						<input name="Sys_DriverID" class="btn1" type="text" value="<%=Ucase(request("Sys_DriverID"))%>" size="10" maxlength="12">
						</td>
						<!--送達狀況
						<select Name="Sys_SendCase" class="btn1">
							<option value="">請選擇</option>
							<option value="1"<%if trim(request("Sys_SendCase"))="1" then response.write " selected"%>>末送達</option>
							<option value="2"<%if trim(request("Sys_SendCase"))="2" then response.write " selected"%>>已送達</option>
						</select>
						<img src="space.gif" width="3" height="10">-->
						<!--<br>
						舉發類型
						<select name="Sys_BillTypeID" class="btn1">
							<option Value="">全部</option>
							<option value="1"<%if trim(request("Sys_BillTypeID"))="1" then response.write " Selected"%>>慢車</option>
							<option value="2"<%if trim(request("Sys_BillTypeID"))="2" then response.write " Selected"%>>行人</option>
							<option value="3"<%if trim(request("Sys_BillTypeID"))="3" then response.write " Selected"%>>道路障礙</option>
						</select>
						</td>-->
						<td>
						列表排序
						</td><td>
						<select Name="Sys_Order" class="btn1">
							<option value="DoubleCheckStatus,BillNo"<%if trim(request("Sys_Order"))="DoubleCheckStatus,BillNo" then response.write " selected"%>>建檔序號</option>
							<option value="Driver,BillNo"<%if trim(request("Sys_Order"))="Driver,BillNo" then response.write " selected"%>>違規人,單號</option>
							<option value="Driver,SENDDATE,BillNo"<%if trim(request("Sys_Order"))="Driver,SENDDATE,BillNo" then response.write " selected"%>>違規人,移送日期,單號</option>
							<option value="IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="IllegaLDate,BillNo" then response.write " selected"%>>違規日期</option>
							<option value="Rule1,Driver,IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="Rule1,Driver,IllegaLDate,BillNo" then response.write " selected"%>>法條,違規人,違規日期</option>
							<option value="Driver,IllegaLDate,Rule1,BillNo"<%if trim(request("Sys_Order"))="Driver,IllegaLDate,Rule1,BillNo" then response.write " selected"%>>違規人,違規日期,法條</option>
							<!--
							<%if showCreditor then%>
								<option value="PetitionDate,Driver,BillNo"<%if trim(request("Sys_Order"))="PetitionDate,Driver,BillNo" then response.write " selected"%>>取得債權日,違規人,單號</option>
							<%end if%>
							-->
							
						</select>
						</td></tr><tr>
						<td>違規日期</td>
						<td nowrap>
							<input name="IllegalDate1" class="btn1" type="text" value="<%=request("IllegalDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate1');">
						~
							<input name="IllegalDate2" class="btn1" type="text" value="<%=request("IllegalDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate2');">
						</td>
						
						
							<td>催告日期</td>
							<td nowrap>
								<input name="UrgeDate1" class="btn1" type="text" value="<%=request("UrgeDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('UrgeDate1');">
							~
								<input name="UrgeDate2" class="btn1" type="text" value="<%=request("UrgeDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('UrgeDate2');">
							</td>
						
						<td>付費日期</td>
						<td nowrap>
							<input name="PayDate1" class="btn1" type="text" value="<%=request("PayDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('PayDate1');">
						~
							<input name="PayDate2" class="btn1" type="text" value="<%=request("PayDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('PayDate2');">
						</td></tr><tr>
						<td>裁決日期</td>
						<td nowrap>
							<input name="JudeDate1" class="btn1" type="text" value="<%=request("JudeDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('JudeDate1');">
						~
							<input name="JudeDate2" class="btn1" type="text" value="<%=request("JudeDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('JudeDate2');">
						</td>
						<td>移送日期</td>
						<td nowrap>
							<input name="SendDate1" class="btn1" type="text" value="<%=request("SendDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('SendDate1');">
						~
							<input name="SendDate2" class="btn1" type="text" value="<%=request("SendDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('SendDate2');">
						</td>
						<td>應到案日期</td>
						<td nowrap>
							<input name="DeallIneDate1" class="btn1" type="text" value="<%=request("DeallIneDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('DeallIneDate1');">
						~
							<input name="DeallIneDate2" class="btn1" type="text" value="<%=request("DeallIneDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('DeallIneDate2');">
						</td>
						</tr>
						<tr>
							<td>建檔日期</td>
							<td nowrap>
								<input name="RecordDate1" class="btn1" type="text" value="<%=request("RecordDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('RecordDate1');">
							~
								<input name="RecordDate2" class="btn1" type="text" value="<%=request("RecordDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('RecordDate2');">
							</td>
							<td>填單日期</td>
							<td nowrap>
								<input name="BillFillDate1" class="btn1" type="text" value="<%=request("BillFillDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('BillFillDate1');">
							~
								<input name="BillFillDate2" class="btn1" type="text" value="<%=request("BillFillDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('BillFillDate2');">
							</td>
							<td>法條代碼</td>
							<td nowrap>
								<input name="Sys_Rule" class="btn1" type="text" value="<%=request("Sys_Rule")%>" size="9">
								
								&nbsp;&nbsp;

								<input type="submit" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funSelt();">
								<input type="button" name="btnCls" value="清除" class="btn3" style="width:40px;height:20px;" onClick="location='PasserBaseQry.asp'">
								<%if trim(Session("Credit_ID"))="A000000000" then%>
									<input type="submit" name="btnSelt" value="上傳" class="btn3" style="width:40px;height:20px;" onclick="uploadFtp();">
								<%end if%>
							</td>
						</tr>
						
						<tr>
							<td>結案日期</td>
							<td nowrap>
								<input name="CaseCloseDate1" class="btn1" type="text" value="<%=request("CaseCloseDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('CaseCloseDate1');">
							~
								<input name="CaseCloseDate2" class="btn1" type="text" value="<%=request("CaseCloseDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('CaseCloseDate2');">
							</td>
							<td>移送案號</td>
							<td>
								<input name="Sys_SendNumber" class="btn1" type="text" value="<%=request("Sys_SendNumber")%>" size="25">
							</td>
							<td>收據號碼</td>
							<td>
								<input name="Sys_PayNo" class="btn1" type="text" value="<%=request("Sys_PayNo")%>" size="25">
							</td>
							<!--
							<td>代保管物</td>
							<td>
								<select Name="Sys_Fastener1" class="btn1">
									<option value="">全部</option><%
										strItem="select * from Code where TypeID=2 and Not(ID<478 or ID=479) order by ID"

										set rsItem=conn.execute(strItem)
										While Not rsItem.Eof
											Response.Write "<option value="""&trim(rsItem("ID"))&""""

											if trim(request("Sys_Fastener1"))=trim(rsItem("ID")) then response.write " selected"

											Response.Write ">"
											Response.Write trim(rsItem("Content"))
											Response.Write "</option>"
											rsItem.MoveNext
										Wend
										rsItem.close
										set rsItem=nothing
									%>
								</select>
							</td>
							-->
						</tr>
						<tr>
							<td>車號</td>
							<td>
								<input name="Sys_CarNo" class="btn1" type="text" value="<%=request("Sys_CarNo")%>" size="15">
							</td>
							<td height="40">專案代碼</td>
							<td nowrap>
								<div id="Layer_ProjectID_ID" style="position:absolute ; z-index:2">
									<input name="Sys_ProjectID" style="width:117px; height:24px;" class="btn1" type="text" value="<%=request("Sys_ProjectID")%>" size="11">
								
								</div>
								<div id="Layer_ProjectID_Selt" style="position:relative ;top:0px;z-index:1">
									<select Name="Sys_ProjectID_Selt" style="width:140px; height:28px;" class="btn1" onchange="myForm.Sys_ProjectID.value=this.value;">
										<option value="">全部</option>
										<% if sys_City="台中市" then %>
										<option value="traffic001','traffic002','traffic003','traffic004','traffic005','traffic006','traffic007','traffic008','traffic009','traffic010">所有慢車</option>
										<% end If %>
										<option value="traffic009">微型電動二輪車</option>
										<option value="traffic004">行人</option>
										<option value="traffic005">道路障礙</option>
										<option value="traffic006">攤販</option>
										<option value="traffic007">人力</option>
										<option value="traffic008">獸力</option>
										<option value="traffic001">自行車</option>
										<option value="traffic002">電動自行車</option>
										<option value="traffic003">電動輔助自行車</option>
										<option value="traffic010">個人行動器具</option>

									</select>
								</div>
							</td>
							<td height="40">批號</td>
							<td>
								<input name="Sys_batchnumber" style="width:117px; height:24px;" class="btn1" type="text" value="<%=UCase(trim(request("Sys_batchnumber")))%>" size="11">
							</td>

						</tr>
						</tr>
						<%If showCreditor then%>
						<tr bgcolor="#CCFFCC">
							<td>再次移送日</td>
							<td nowrap>
								<input name="Sys_SendDetailDate1" class="btn1" type="text" value="<%=request("Sys_SendDetailDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_SendDetailDate1');">
							~
								<input name="Sys_SendDetailDate2" class="btn1" type="text" value="<%=request("Sys_SendDetailDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_SendDetailDate2');">
							</td>

							<td>債權狀況</td>
							<td>
								<select name="sys_CreditorTypeID">

									<option value="-1"<%if trim(Request("sys_CreditorTypeID"))="-1" then response.write " Selected"%>>未申請債權</option>

									<option value=""<%if trim(Request("sys_CreditorTypeID"))="" then response.write " Selected"%>>全部</option>

									<option value="0','1"<%if trim(Request("sys_CreditorTypeID"))="0','1" then response.write " Selected"%>>已申請債權</option>

									<option value="0"<%if trim(Request("sys_CreditorTypeID"))="0" then response.write " Selected"%>>清償中</option>

									<option value="1"<%if trim(Request("sys_CreditorTypeID"))="1" then response.write " Selected"%>>無個人財產</option>

								</select>
							</td>

							<td>確定日期</td>
							<td nowrap>
								<input name="MakeSureDate1" class="btn1" type="text" value="<%=request("MakeSureDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('MakeSureDate1');">
							~
								<input name="MakeSureDate2" class="btn1" type="text" value="<%=request("MakeSureDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('MakeSureDate2');">
							</td>
						</tr>
						<tr bgcolor="#CCFFCC">
							<td>債權取得日(1)</td>
							<td nowrap>
								<input name="Sys_PetitionDate1" class="btn1" type="text" value="<%=request("Sys_PetitionDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_PetitionDate1');">
							~
								<input name="Sys_PetitionDate2" class="btn1" type="text" value="<%=request("Sys_PetitionDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_PetitionDate2');">
							</td>
							<td>債權取得日(N)</td>
							<td nowrap>
								<input name="max_PetitionDate1" class="btn1" type="text" value="<%=request("max_PetitionDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('max_PetitionDate1');">
							~
								<input name="max_PetitionDate2" class="btn1" type="text" value="<%=request("max_PetitionDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('max_PetitionDate2');">
							</td><!--
							<td>債權取得狀況</td>
							<td nowrap>
								<select Name="Sys_CreditorType" class="btn1">
									<option value="">全部</option>
									<option value="1"<%if trim(request("Sys_CreditorType"))="1" then response.write " selected"%>>未取得</option>
							
								</select>
							</td>-->
							<td>外國人</td>
							<td nowrap>
								<select Name="Sys_DriverType" class="btn1">
									<option value="">全部</option>
									<option value="1"<%if trim(request("Sys_DriverType"))="1" then response.write " selected"%>>是</option>
									<option value="1"<%if trim(request("Sys_DriverType"))="2" then response.write " selected"%>>否</option>
							
								</select>
							</td>
						</tr>
						<tr bgcolor="#CCFFCC">							
							<td>微電車大宗號碼</td>
							<td colspan="8">
								<input name="Sys_SendMailStation" class="btn1" type="text" value="<%=request("Sys_SendMailStation")%>" size="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							</td>
						</tr>
							<%if sys_City="彰化縣" or sys_City="基隆市" then%>
								<tr bgcolor="#CCFFCC">
									<td>保留年度</td>
									<td>
										<input name="Sys_ReserveYear1" class="btn1" type="text" value="<%=request("Sys_ReserveYear1")%>" size="1" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
										&nbsp;&nbsp;~&nbsp;&nbsp;
										<input name="Sys_ReserveYear2" class="btn1" type="text" value="<%=request("Sys_ReserveYear2")%>" size="1" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">

									</td>
									<%if sys_City="彰化縣" then %>
										<td>清查年度</td>
										<td colspan=3>
											<input name="Sys_checkdate1" class="btn1" type="text" value="<%=request("Sys_checkdate1")%>" size="1" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
											&nbsp;&nbsp;~&nbsp;&nbsp;
											<input name="Sys_checkdate2" class="btn1" type="text" value="<%=request("Sys_checkdate2")%>" size="1" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">

										</td>
									<%end if%>
								</tr>

							<%end if%>
						<%end if%>
						</table>
						<!--<hr>
						催告日期
						<input name="Sys_sltUrgeDate1" class="btn1" type="text" value="<%=request("Sys_sltUrgeDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltUrgeDate1');">
						<!-- ~
						<input name="Sys_sltUrgeDate2" class="btn1" type="text" value="<%=request("Sys_sltUrgeDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltUrgeDate2');">-->
						<!--<input type="button" name="btnSelt" value="查詢" onclick="funSltUrgeDate();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(224,1)=false then
							response.write "disabled"
						end if
						%>>
						裁決日期
						<input name="Sys_sltJudeDate1" class="btn1" type="text" value="<%=request("Sys_sltJudeDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltJudeDate1');">
						<!-- ~
						<input name="Sys_sltJudeDate2" class="btn1" type="text" value="<%=request("Sys_sltJudeDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltJudeDate2');">-->
						<!--<input type="button" name="btnSelt" value="查詢" onclick="funSltJudeDate();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(224,1)=false then
							response.write "disabled"
						end if
						%>>
						移送日期
						<input name="Sys_sltSendDate1" class="btn1" type="text" value="<%=request("Sys_sltSendDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltSendDate1');">
						<!-- ~
						<input name="Sys_sltSendDate2" class="btn1" type="text" value="<%=request("Sys_sltSendDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_sltSendDate2');">-->
						<!--<input type="button" name="btnSelt" value="查詢" onclick="funSltSendDate();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(224,1)=false then
							response.write "disabled"
						end if
						%>>-->
						<HR>
							未做裁決且未繳費
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funUrgeDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							裁決後
							未繳費且未催告
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funJudeDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							裁決後
							未繳費且未移送
							<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funSendDateSelt();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,1)=false then
								response.write "disabled"
							end if
							%>>
							
							<%if sys_City="彰化縣" then%>
								當年度清查案件
								<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funPasserProperty();">

							<%end if%>

							<input type="button" name="btnSelt" class="btn3" style="width:80px;height:20px;" value="資料回復" onclick="funRecallData();" <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(224,3)=false then
								response.write "disabled"
							end if
							%>>
							<%'if sys_City="基隆市" or sys_City="高雄市" or sys_City="台中市" then%>
							&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
							<input type="button" name="btnSelt" value="整批戶籍地址更正" onclick="funUpdAddress();" >
							<%'end if%>
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
					<th class="font10" nowrap>車號</th>
					<!--<th class="font10" nowrap>舉發人</th>-->
					<th class="font10" nowrap>違規人</th>
					<th class="font10" nowrap>法條</th>
					<th class="font10" nowrap>裁決日</th>
					<%'if sys_City<>"基隆市" then%>
						<th class="font10" nowrap>催告日</th>
					<%'end if%>
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
					<% If sys_City="彰化縣" then %>
						<th class="font10" nowrap>清查日</th>
					<% end If %>
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
					if rsfound.eof then exit For 
					
					response.write "<tr align='center' bgcolor='#FFFFFF'"
					lightbarstyle 0
					response.write ">"
					response.write "<td class=""font10""><input class=""btn1"" type=""checkbox"" name=""chkSend"" value="""&trim(rsfound("Sn"))&""" onclick=funChkSend();></td>"
					response.write "<td class=""font10"">"&trim(rsfound("DoubleCheckStatus"))&"</td>"
					response.write "<td class=""font10"">"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("BillNo"))
					If trim(rsfound("CARSIMPLEID")) = "8" Then

						Sql_ChkDCI="select count(1) cnt from PASSERDCILOG where exchangetypeid='W' and billsn="&trim(rsfound("Sn"))

						set rsdci=conn.execute(Sql_ChkDCI)
						If cdbl(rsdci("cnt")) = 0 Then Response.Write "<br><span class=""err01"">入案未上傳</span>"
						rsdci.close

						Sql_ChkDCI="select count(1) cnt from PASSERDCILOG where exchangetypeid='W'  and dcireturnstatusid is null and billsn="&trim(rsfound("Sn"))

						set rsdci=conn.execute(Sql_ChkDCI)
						If cdbl(rsdci("cnt")) >0 Then Response.Write "<br><span class=""err01"">入案未處理</span>"
						rsdci.close

						Sql_ChkDCI="select count(1) cnt from PASSERDCILOG where exchangetypeid='W'  and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid='W' and dcireturnstatus=-1) and billsn="&trim(rsfound("Sn"))

						set rsdci=conn.execute(Sql_ChkDCI)
						If cdbl(rsdci("cnt")) > 0 Then Response.Write "<br><span class=""err01"">入案異常</span>"
						rsdci.close
					
					End if 
					
					Response.Write "</td>"
					response.write "<td class=""font10"">"&trim(rsfound("Carno"))&"</td>"
					'response.write "<td class=""font10"">"&trim(rsfound("BillMem1"))&"</td>"
					'response.write "<td><a href='../BillKeyIn/BillKeyIn_People_Update.asp?BillSN="&trim(rsfound("SN"))&"' target='_blank'>"&trim(rsfound("BillNo"))&"</a></td>"

					response.write "<td class=""font10"">"&trim(rsfound("Driver"))&"</td>"
					'response.write "<td>"&trim(rsfound("IllegalAddress"))&"</td>"

					if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")					
					if rsfound("Rule2")<>"" then chRule=chRule&"\"&rsfound("Rule2")
					
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
					response.write "</td>"
					'response.write "<td>"&FORFEIT&"</td>"
					'smith 20091005 基隆不用秀
					response.write "<td class=""font10"">"&trim(gInitDT(rsfound("JUDEDATE")))&"</td>"

					'if sys_City<>"基隆市" then
						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("URGEDATE")))&"</td>"
					'End if 

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

						strSD="select min(PetitionDate) PetitionDate from PasserCreditor where  BillSn="&Trim(rsfound("sn"))

						Set rsSD=conn.execute(strSD)

						If Not rsSD.eof then	
							SENDDATEtmp=Trim(rsSD("PetitionDate"))
						End If
						rsSD.close
						Set rsSD=Nothing 

						response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"

						SENDDATEtmp2=""

						strSD="select max(PetitionDate) PetitionDate from PasserCreditor where  BillSn="&Trim(rsfound("sn"))&" and (select count(1) cnt from PasserCreditor pc1 where BillSn="&Trim(rsfound("sn"))&")>1"

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
					End If 

					if sys_City="彰化縣" then

						If (DB_KindSelt="PasserProperty" or ( request("Sys_checkdate1")<>"" and request("Sys_checkdate2")<>"") ) and show_postNumber="" then

							show_postNumber=rsfound("postNumber")
						end If 

						response.write "<td class=""font10"">"&trim(gInitDT(rsfound("checkdate")))&"</td>"
					end If 

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

						If trim(rsfound("CARSIMPLEID")) = "8" Then

							Sql_ChkDCI="select count(1) cnt from PasserDCILog where ReturnMarkType=9 and sn="&trim(rsfound("dcilogsn"))&" and billsn="&trim(rsfound("Sn"))

							set rsdci=conn.execute(Sql_ChkDCI)
							If cdbl(rsdci("cnt")) = 0 Then Response.Write "<br><span class=""err01"">結案未上傳</span>"
							rsdci.close

							Sql_ChkDCI="select count(1) cnt from PasserDCILog where ReturnMarkType=9 and sn="&trim(rsfound("dcilogsn"))&" and billsn="&trim(rsfound("Sn"))&" and dcireturnstatusid is null"

							set rsdci=conn.execute(Sql_ChkDCI)
							If cdbl(rsdci("cnt")) = 1 Then Response.Write "<br><span class=""err01"">結案未處理</span>"
							rsdci.close

							

							Sql_ChkDCI="select count(1) cnt from PasserDCILog where ReturnMarkType=9 and sn="&trim(rsfound("dcilogsn"))&" and billsn="&trim(rsfound("Sn"))&" and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid ='N' and dcireturnstatus=-1)"

							set rsdci=conn.execute(Sql_ChkDCI)
							If cdbl(rsdci("cnt")) = 1 Then Response.Write "<br><span class=""err01"">結案異常</span>"
							rsdci.close
						
						End if 
					end if
					response.write "</td>"
%>					<td class="font10" nowrap align="left">

						<input type='button' value='繳款' class="btn3" style="width:40px;height:20px;" onclick='window.open("<%
							if sys_City="彰化縣" then

								Response.Write "Passer_Pay_Sys.asp?PBillSN="&trim(rsfound("SN"))
							else

								Response.Write "Passer_Pay.asp?PBillSN="&trim(rsfound("SN"))
							End if 
						%>","WebPage4","left=0,top=0,location=0,width=1100,height=575,resizable=yes,scrollbars=yes")'>
							
							<input type='button' value='裁決' class="btn3" style="width:40px;height:20px;" onclick='window.open("Passer_Jude.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage1","left=0,top=0,location=0,width=950,height=700,status=yes,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>	
							
							<input type='button' value='送達' class="btn3" style="width:40px;height:20px;" onclick='window.open("PasserSendArrived.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=1100,height=575,resizable=yes,scrollbars=yes")'>

							<input type='button' value='催告' class="btn3" style="width:40px;height:20px;" onclick='window.open("PasserUrgeDetail.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage2","left=0,top=0,location=0,width=1000,height=600,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>
							
						
						<input type='button' value='移送' class="btn3" style="width:40px;height:20px;" onclick='window.open("PasserSendDetail.asp?PBillSN=<%=trim(rsfound("SN"))%>","WebPage3","left=0,top=0,location=0,width=1000,height=600,resizable=yes,scrollbars=yes")' <%if trim(rsfound("BILLSTATUS"))="9" and sys_City="宜蘭縣" then Response.Write "disabled"%>>

						<%
						if sys_City="基隆市" then
						%>
						<input type='button' value='收文' class="btn3" style="width:40px;height:20px;" onclick='window.open("PasserEtaxOption.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=900,height=575,resizable=yes,scrollbars=yes")'>
						<%
						end If 
						
						%>
					
						<!--<input type='button' value='執行處回文' onclick='window.open("Passer_Send.asp","WebPage2","left=0,top=0,location=0,width=500,height=455,resizable=yes,scrollbars=yes")'>
						
						<br>-->						

						<%if showCreditor then%>
							<input type='button' value='債權' class="btn3" style="width:40px;height:20px;" onclick='window.open("PasserCreditor.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=900,height=575,resizable=yes,scrollbars=yes")'<%If Ifnull(rsfound("SendDate")) Then Response.Write " disabled"%>>
						<%end If 
						
						if trim(rsfound("BILLSTATUS"))="9" then
							If trim(rsfound("CARSIMPLEID")) = "8" Then
						%>

							<input type='button' value='撤消' class="btn3" style="width:40px;height:20px;" onclick='CancelFile(<%=trim(rsfound("SN"))%>);'>
						<%
							end If 
						End if 
						%>

						<input type="button" name="Update" value="詳細" class="btn3" style="width:40px;height:20px;" onclick='window.open("../Query/ViewBillBaseData_people.asp?BillSn=<%=trim(rsfound("SN"))%>","WebPage1","left=0,top=0,location=0,width=850,height=700,status=yes,resizable=yes,scrollbars=yes")'>
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

		<td height="35" bgcolor="#1BF5FF" align="center" nowrap>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);" class="btn3" style="width:60px;height:30px;font-size:14px;">
			<span class="style2"><%=fix(cdbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(cdbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);" class="btn3" style="width:60px;height:30px;font-size:14px;">
			<img src="space.gif" width="18" height="8">
			
			<%
				if sys_City="彰化縣" then

					Response.Write "<br>"

					Response.Write "<input name=""Sys_PostNumber"" class=""btn1"" type=""text"" value="""&show_postNumber&""" size=""20"" maxlength=""30"">&nbsp;&nbsp;"
					
					Response.Write "<input type=""button"" name=""btnSaveNumber"" value=""儲存文號"" class=""btn3"" style=""width:120px;height:30px;font-size:15px;"" onClick=""funSetPostNumber();"">"

				end if

			%>

			<input type="button" name="btnExecel" value="轉換成Excel" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funchgExecel();">
			<input type="button" name="btnExecel" value="郵局大宗函件" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funPasserMailMoney();">
			<br>
			
			<input type="button" name="btnExecel" value="批次裁決通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeBat();">

			<% '基隆催告的都不用秀 smith 20091005 
			'if sys_City<>"基隆市" then %>
				<input type="button" name="btnExecel" value="批次催繳通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funUrgeBat();">
			<%' end if %>

			<input type="button" name="btnExecel" value="批次移送通知" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funSendBat();">

			<input type="button" name="btnExecel" value="行政執行移送電子清冊" class="btn3" style="width:160px;height:30px;font-size:14px;" onclick="funCountryBat();">
			<%if showCreditor then%>
				<input type="button" name="btnExecel" class="btn3" style="width:120px;height:30px;font-size:14px;" value="批次債權移送" onclick="funSendBatTwo_chromat();">
			<%end if%>

			<% if sys_City="基隆市" then %>
				<br><input type="button" name="btnETAX" value="批次國稅局收文登記" class="btn3" style="width:270px;height:30px;font-size:14px;" onclick="funPasserEtax()">
			<% end if %>

			<% if sys_City="台中市" then %>
					<input type="button" name="btnExecel" value="批次送達註記" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funArrivedBat();">
			<% end if %>

			<% If trim(Session("Credit_ID"))="A000000000" then %>
			<!--
				<input type="button" name="btnExecel" value="批次債權註記" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funCreditorBat();">

				
				<input type="button" name="btnExecel" value="批次債權掃描檔上傳" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funCreditorimg();">
-->
			<% end If %>

			<%if sys_City="台中市" or sys_City="台中縣" then%>
				<input type="button" name="btnAllPrint" value="批次舉發單列印" class="btn3" style="width:150px;height:30px;font-size:14px;" onclick="PasserAllPrint()">


			<!--
			<br>
			<input type="button" name="btnExecel" value="批次催繳套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funUrgeBat_chromat();">
			
			<input type="button" name="btnExecel" value="批次裁決套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeBat_chromat();">

			<input type="button" name="btnExecel" value="批次移送套印" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funSendBat_chromat();">
			-->
			<%end if%>

			
		</td>
	</tr>
	
	<tr>		
		<td height="35" bgcolor="#FFFFFF" align="center" nowrap>
			<!-- <font size="2">依據上方選擇案件資料產生相關清冊</font> -->
			<span class="style3"><img src="space.gif" width="8" height="8"></span>

			<input type="button" name="Submite32" value="建檔清冊" class="btn3" style="width:80px;height:30px;font-size:14px;" onclick="funPrintPeopleList_Stop();">

			<input type="button" name="btnExecel" value="舉發清冊(xls)" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeExecel();">

			
			<% if sys_City="台南市" or sys_City="基隆市" then %>

				<input type="button" name="btnExecel" value="慢車註銷案件清冊" class="btn3" style="width:160px;height:30px;font-size:14px;" onclick="funWriteOff();">

				<input type="button" name="btnExecel" value="應收款項註銷清冊" class="btn3" style="width:160px;height:30px;font-size:14px;" onclick="funWriteOffOption();">
			
			<% End if %>
			<br>

			<input type="button" name="btnExecel" value="裁決清冊(xls)" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funJudeListExecel()">

			<input type="button" name="btnExecel" value="裁決公示送達清冊" class="btn3" style="width:130px;height:30px;font-size:14px;" onclick="funGovArriveBat();">

			<input type="button" name="btnExecel" value="裁決寄存送達清冊" class="btn3" style="width:130px;height:30px;font-size:14px;" onclick="funSendArriveBat();">

			<%if showCreditor then 'PasserPayDetail_YiLanList.asp%>
				<br>
				<input type="button" name="btnExecel" class="btn3" style="width:100px;height:30px;font-size:14px;" value="未送達清冊" onclick="funPassersEndArrivedList();">
				<input type="button" name="btnExece2" class="btn3" style="width:130px;height:30px;font-size:14px;" value="未繳納待執行清冊" onclick="funNoSendList();">
				<input type="button" name="btnExece3" class="btn3" style="width:160px;height:30px;font-size:14px;" value="已移送未執行債權清冊" onclick="funPasserSendnotCredit();">
				<input type="button" name="btnExece3" class="btn3" style="width:170px;height:30px;font-size:14px;" value="債權憑證準備再移送清冊" onclick="funPasserSendtwoList();">
				<input type="button" name="btnExece3" class="btn3" style="width:150px;height:30px;font-size:14px;" value="交付保管品核對清冊" onclick="funPasserInventoryList();">
				<input type="button" name="btnExece3" class="btn3" style="width:150px;height:30px;font-size:14px;" value="行政罰鍰收繳情形明細表" onclick="funPasserPayDetail_YiLanList();">
				<input type="button" name="btnExece3" class="btn3" style="width:150px;height:30px;font-size:14px;" value="應收未收收繳情形明細表" onclick="funPasserNoPayDetail_YiLanList();">
				<br>
				<input type="button" name="btnExecel" class="btn3" style="width:120px;height:30px;font-size:14px;" value="無個人財產清冊" onclick="funNoEffectsList();">

				<input type="button" name="btnExecel" value="債權憑證清冊" class="btn3" style="width:100px;height:30px;font-size:14px;" onclick="funCreditorList();">

				
				<input type="button" name="btnExecel" value="債權執行狀態清冊" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funCreditorStstusList();">

				<input type="button" name="btnExecel" value="債權明細統計表" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funPasserSendCreditorList();">

				<input type="button" name="btnExecel" value="債權再移送明細統計表" class="btn3" style="width:160px;height:30px;font-size:14px;" onclick="funPasserSendCreditorTwoList();">
			<%end if

			'基隆催告的都不用秀 smith 20091005
			if sys_City<>"基隆市" then %>
				<br>
				<input type="button" name="btnExecel" value="催告清冊" class="btn3" style="width:70px;height:30px;font-size:14px;" onclick="funSendExecel()">

				<input type="button" name="btnExecel" value="催告已到案清冊" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funSendCloseExecel();">

				<input type="button" name="btnExecel" value="催告未到案清冊" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funUcloseExecel();">
			<% end if %>
			<br>

			<input type="button" name="Submite32" value="移送案件明細表" class="btn3" style="width:110px;height:30px;font-size:14px;" onclick="funPasserSend_Stop();">

			<input type="button" name="Submite32" value="繳費明細表" class="btn3" style="width:90px;height:30px;font-size:14px;" onclick="funPasserPay_Stop();">

			<input type="button" name="btnExecel" value="收繳費統計表" class="btn3" style="width:100px;height:30px;font-size:14px;" onclick="funUnitListExecel();">

			<input type="button" name="Submite32" value="未繳費明細表" class="btn3" style="width:100px;height:30px;font-size:14px;" onclick="funPasserNoPay_Stop();">
			
			<input type="button" name="btnExecel" value="未繳費統計表" class="btn3" style="width:100px;height:30px;font-size:14px;" onclick="funNoPayUnitListExecel();">			
			
			<input type="button" name="Submite32" value="年度保留清冊" class="btn3" style="width:100px;height:30px;font-size:14px;" onclick="funPasser_Reserve();">
			
			<% if sys_City="彰化縣" then %>
				<br><input type="button" name="btnExecel" value="(署)處理違反道路交通管理事件統計表" class="btn3" style="width:270px;height:30px;font-size:14px;" onclick="funPayReport()">	
				
				<input type="button" name="btnExecel" value="交通罰緩收入憑證月報表" class="btn3" style="width:190px;height:30px;font-size:14px;" onclick="funPayReportMonth()">	
				
				<input type="button" name="btnExecel" value="(分局)交通罰緩收入憑證月報表" class="btn3" style="width:210px;height:30px;font-size:14px;" onclick="funPayReportMonth_Unit()">

				<input type="button" name="btnExecel" value="(分局)交通罰緩收據明細表" class="btn3" style="width:190px;height:30px;font-size:14px;" onclick="funPayReportDay_Unit()">

				<br>

				<input type="button" name="btnIRSList" value="國稅局清冊" class="btn3" style="width:190px;height:30px;font-size:14px;" onclick="funIRSList()" <%if DB_KindSelt<>"PasserProperty" then Response.Write " disabled "%>>

				<input type="button" name="btnTaxList" value="稅務局清冊" class="btn3" style="width:190px;height:30px;font-size:14px;" onclick="funTaxList()" <%if DB_KindSelt<>"PasserProperty" then Response.Write " disabled "%>>

				<input type="button" name="btnExecel" value="(分局)收納款項收據紀錄卡" class="btn3" style="width:190px;height:30px;font-size:14px;" onclick="funPayReport_UnitList()">
				<br>
				<input type="button" name="btnExecel" value="交通違規案件行政罰鍰憑證清冊" class="btn3" style="width:250px;height:30px;font-size:14px;" onclick="funPasserBaseSendCreditor()">
				&nbsp;&nbsp;
				<input type="button" name="btnExecel" value="交通違規案件行政罰鍰第一次憑證清冊" class="btn3" style="width:270px;height:30px;font-size:14px;" onclick="funPasserBaseSendCreditorOne()">
			<% end if %>

			<% if sys_City="台中市" then %>
				<br><input type="button" name="btnExecel" value="各單位清理統計表" class="btn3" style="width:270px;height:30px;font-size:14px;" onclick="funPayReport2()">
				<input type="button" name="btnExece2" value="各單位清理進度表" class="btn3" style="width:270px;height:30px;font-size:14px;" onclick="funPayReport3()">

				<input type="button" name="btnExecel" value="交通違規罰鍰繳款明細表" class="btn3" style="width:200px;height:30px;font-size:14px;" onclick="funPasserPayCloseDetail()">

			<% end if %>
				<input type="button" name="btnExecel" value="微電車清冊" class="btn3" style="width:200px;height:30px;font-size:14px;" onclick="funMinuteCarList()">
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
<input type="Hidden" name="Sys_Execution" value="">
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
<input type="Hidden" name="hd_CanCelSN" value="">
<center><B>
<font size="6" color="red">請先至『單位管理系統』填寫單位相關資料。</font>
</B></center>
</form>

<form name="exForm" method="post">
	<input type="Hidden" name="orderstr" value="<%=orderstr%>">
	<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
	<input type="Hidden" name="Sys_SendBillSN" value="">
	<input type="Hidden" name="hd_BillSN" value="">
	<input type="Hidden" name="ExportSQL" value="<%=tmpSQL%>">
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

function PasserAllPrint(){
	UrlStr="../PasserQuery/PasserAllPrint.asp";
	myForm.action=UrlStr;
	myForm.target="PasserAllPrint";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function uploadFtp(){
	UrlStr="PasserBaseFtpOpen.asp";
	myForm.action=UrlStr;
	myForm.target="UpFtp";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function CancelFile(SN){
	if(confirm('確定要上傳取消結案嗎？')){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="Selt";
		myForm.DB_Display.value='show';
		myForm.Sys_SendBillSN.value='';
		myForm.DB_KindSelt.value="Cancel";
		myForm.hd_CanCelSN.value=SN;
		myForm.submit();
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

function funPayReport_UnitList(){
	myForm.action="ReportMonthPay_Unit_List.asp";
	myForm.target="funPayReport_UnitList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
}

function funPasserBaseSendCreditor(){
	myForm.action="PasserBaseSendCreditorList_Execel.asp";
	myForm.target="PasserBaseSendCreditor";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
}

function funPasserBaseSendCreditorOne(){
	myForm.action="PasserBaseSendCreditorList_One_Execel.asp";

	myForm.target="funPasserBaseSendCreditorOne";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
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


function funIRSList(){
	
	myForm.action="PasserIRSList.asp";
	myForm.target="IRSList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
}

function funTaxList(){
	
	myForm.action="PasserTaxList.asp";
	myForm.target="TaxList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
		
}


function funPayReport(){

	myForm.action="../Report/Report0110.asp";
	myForm.target="funPayReport";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}



function funMinuteCarList(){
	
	myForm.action="PasserMinuteCarList.asp";
	myForm.target="MinuteCarList";
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


function funPasserEtax(){
	
	myForm.action="PasserEtax.asp";
	myForm.target="PasserEtax";
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

function funPasserProperty(){
	var error=0;
	myForm.hd_BillSN.value='';
	myForm.hd_BillNo.value='';
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="PasserProperty";
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

function funWriteOff(){
	if(myForm.DB_Display.value=="show"){
		//newWin("","PasserBaseQry_Execel",980,550,0,0,"yes","yes","yes","no");
		UrlStr="PasserBaseWriteOff_TaiNan.asp";
		myForm.action=UrlStr;
		myForm.target="PasserBaseWriteOff_TaiNan";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funWriteOffOption(){
	if(myForm.DB_Display.value=="show"){
		//newWin("","PasserBaseQry_Execel",980,550,0,0,"yes","yes","yes","no");
		UrlStr="PasserBaseWriteOffOption.asp";
		myForm.action=UrlStr;
		myForm.target="PasserBaseWriteOffOption";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funchgExecel(){
	if(myForm.DB_Display.value=="show"){
		//newWin("","PasserBaseQry_Execel",980,550,0,0,"yes","yes","yes","no");
		
		if (exForm.DB_Cnt.value<=5000){

			exForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
			exForm.hd_BillSN.value=myForm.hd_BillSN.value;
			exForm.ExportSQL.value="";

		}else{			
			exForm.Sys_SendBillSN.value="";
			exForm.hd_BillSN.value="";
		}

		UrlStr="PasserBaseQry_Execel.asp";
		exForm.action=UrlStr;
		exForm.target="PasserBaseQry_Execel";
		exForm.submit();
		exForm.action="";
		exForm.target="";
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


function funSetPostNumber(){
	if(myForm.DB_KindSelt.value=="PasserProperty" || ((myForm.Sys_SendBillSN.value!='' || myForm.hd_BillSN.value!='') && myForm.Sys_checkdate1.value!='' && myForm.Sys_checkdate2.value!='')){

		myForm.action="PasserSetPostNumber.asp";
		myForm.target="SetPostNumber";
		myForm.submit();
		myForm.action="";
		myForm.target="";
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
function funArrivedBat(){
		exForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		exForm.hd_BillSN.value=myForm.hd_BillSN.value;

		myForm.action="PasserSendArrivedBat.asp";
		myForm.target="PasserSendArrivedBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funCreditorBat(){
		exForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		exForm.hd_BillSN.value=myForm.hd_BillSN.value;

		myForm.action="PasserCreditor_Bat.asp";
		myForm.target="PasserCreditor_Bat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}



function funCreditorimg(){
		exForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		exForm.hd_BillSN.value=myForm.hd_BillSN.value;

		myForm.action="PasserCreditor_SendStyle.asp";
		myForm.target="SendStyle_Creditor";
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
		myForm.Sys_Execution.value="";
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
		myForm.Sys_Execution.value="";
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
		myForm.Sys_Execution.value="";
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
		myForm.Sys_Execution.value="";
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
		myForm.Sys_Execution.value="";
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

function funMap(SN){
	UrlStr="SendStyle_PasserImage.asp?SN="+SN;
	newWin(UrlStr,"winMap",700,550,50,10,"yes","yes","yes","no");
}

</script>
<%
conn.close
set conn=nothing
%>