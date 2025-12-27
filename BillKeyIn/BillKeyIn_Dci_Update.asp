<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<title>舉發單修改</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)
	
	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
end function

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 
'==========================
'修改告發單
if trim(request("kinds"))="DB_insert" then
	if trim(request("Billno1"))<>trim(request("OldBillNo")) then
		strUpd1="Update BillBase set BillNo='"&trim(request("Billno1"))&"' where SN="&trim(request("BillSN"))
		conn.execute strUpd1
		strUpd2="Update DciLog set BillNo='"&trim(request("Billno1"))&"' where BillSn="&trim(request("BillSN"))
		conn.execute strUpd2
		strUpd3="Update BillMailHistory set BillNo='"&trim(request("Billno1"))&"' where BillSn="&trim(request("BillSN"))
		conn.execute strUpd3
		'如果修改單號、車號，會造成資料與監理站資料不符
		strUpd4="Update BillBaseDciReturn  set BillNo='"&trim(request("Billno1"))&"' where BillNo='"&trim(request("OldBillNo"))&"'"
		conn.execute strUpd4
	end if
	if trim(request("NewCarNo"))<>trim(request("OldCarNo")) then
		strUpd1="Update BillBase set CarNo='"&trim(request("NewCarNo"))&"' where SN="&trim(request("BillSN"))
		conn.execute strUpd1
		strUpd2="Update DciLog set CarNo='"&trim(request("NewCarNo"))&"' where BillSn="&trim(request("BillSN"))
		conn.execute strUpd2
		strUpd3="Update BillMailHistory set CarNo='"&trim(request("NewCarNo"))&"' where BillSn="&trim(request("BillSN"))
		conn.execute strUpd3
		strUpd4="Update BillBaseDciReturn  set CarNo='"&trim(request("NewCarNo"))&"' where BillNo='"&trim(request("Billno1"))&"' and CarNo='"&trim(request("OldCarNo"))&"' and ExchangeTypeID<>'A'"
		conn.execute strUpd4
	end if
	'駕駛人生日
	theDriverBirth=""
	if trim(request("BillDriverBirth"))<>"" then
		theDriverBirth=DateFormatChange(trim(request("BillDriverBirth")))
	else 
		theDriverBirth = "null"
	end if
	'BillBase
	strBillUpd="update BillBase set " &_
		"DriverBirth="&theDriverBirth&",DriverID='"&trim(request("BillDriverID"))&"'" &_
		",DriverSex='"&trim(request("BillDriverSex"))&"',Note='"&trim(request("Note"))&"'" &_
		" where SN="&trim(request("BillSN"))

		conn.execute strBillUpd
		ConnExecute strBillUpd,353
%>
<script language="JavaScript">
	alert("舉發單駕駛人資料修改完成");
</script>
<%
end if

'修改DCI回傳資料
if trim(request("kinds"))="DciW_Update" then

	if trim(request("WOwnerAddress"))<>"" then
		strWOZip="select * from Zip where ZipName like '"&left(replace(trim(request("WOwnerAddress")),"臺","台"),5)&"%'"
		set rsWOZip=conn.execute(strWOZip)
		if not rsWOZip.eof then
			updWOwnerZip=trim(rsWOZip("ZipID"))
		else
			if trim(request("WOwnerZip"))<>"" then
				updWOwnerZip=trim(request("WOwnerZip"))
			else
				updWOwnerZip=""
			end if
		end if
		rsWOZip.close
		set rsWOZip=nothing
	end if

	if trim(request("WDriverHomeAddress"))<>"" then
		strWDZip="select * from Zip where ZipName like '"&left(replace(trim(request("WDriverHomeAddress")),"臺","台"),5)&"%'"
		set rsWDZip=conn.execute(strWDZip)
		if not rsWDZip.eof then
			updWDriverZip=trim(rsWDZip("ZipID"))
		else
			if trim(request("WDriverHomeZip"))<>"" then
				updWDriverZip=trim(request("WDriverHomeZip"))
			else	
				updWDriverZip=""
			end if
		end if
		rsWDZip.close
		set rsWDZip=nothing
	end if
	'if sys_City="花蓮縣" Or sys_City="高雄縣" Or sys_City="高雄市" Or sys_City="南投縣" Or sys_City="台東縣" Or sys_City="屏東縣" Or sys_City="台南市" or (sys_City="苗栗縣" And Trim(request("sys_BillTypeID"))="2") then
		strDciWUpdateA="update BillBaseDciReturn set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			",Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverHomeZip='"&updWDriverZip&"'" &_
			",DriverHomeAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where CarNo='"&trim(request("WCarNo"))&"' and ExchangeTypeID='A'" 
		conn.execute strDciWUpdateA
	'end If
	'If sys_City="高雄市" or sys_City="花蓮縣" Or sys_City="南投縣" Or sys_City="屏東縣" Or sys_City="台南市" or sys_City="台中市" or sys_City="保二總隊三大隊一中隊" Then
		strUpdBillBase="Update BillBase set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			",Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverZip='"&updWDriverZip&"'" &_
			",DriverAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where SN="&trim(request("BillSN"))
		conn.execute strUpdBillBase
	'End If 
		strDciWUpdate="update BillBaseDciReturn set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			",Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverHomeZip='"&updWDriverZip&"'" &_
			",DriverHomeAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where BillNo='"&trim(request("WBillNo"))&"' and CarNo='"&trim(request("WCarNo"))&"'" 
	
		conn.execute strDciWUpdate
		ConnExecute strDciWUpdate,353
%>
<script language="JavaScript">
	alert("監理所回傳資料修改完成");
</script>
<%
end if

'修改DCI回傳資料(車主)
if trim(request("kinds"))="DciW_Update1" then

	if trim(request("WOwnerAddress"))<>"" then
		strWOZip="select * from Zip where ZipName like '"&left(replace(trim(request("WOwnerAddress")),"臺","台"),5)&"%'"
		set rsWOZip=conn.execute(strWOZip)
		if not rsWOZip.eof then
			updWOwnerZip=trim(rsWOZip("ZipID"))
		else
			if trim(request("WOwnerZip"))<>"" then
				updWOwnerZip=trim(request("WOwnerZip"))
			else
				updWOwnerZip=""
			end if
		end if
		rsWOZip.close
		set rsWOZip=nothing
	end if

	'if sys_City="花蓮縣" Or sys_City="高雄縣" Or sys_City="高雄市" Or sys_City="南投縣" Or sys_City="屏東縣" Or sys_City="台南市" or (sys_City="苗栗縣" And Trim(request("sys_BillTypeID"))="2") then
		strDciWUpdateA="update BillBaseDciReturn set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			" where CarNo='"&trim(request("WCarNo"))&"' and ExchangeTypeID='A'" 
		conn.execute strDciWUpdateA
	'end If
	'If sys_City="高雄市" or sys_City="花蓮縣" Or sys_City="南投縣" Or sys_City="台南市" Or sys_City="屏東縣" or sys_City="台中市" or sys_City="保二總隊三大隊一中隊" Then
		strUpdBillBase="Update BillBase set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			" where SN="&trim(request("BillSN"))
		conn.execute strUpdBillBase
	'End If 
		strDciWUpdate="update BillBaseDciReturn set " &_
			"Owner='"&trim(request("WOwner"))&"',OwnerID='"&trim(request("WOwnerID"))&"'" &_
			",OwnerZip='"&updWOwnerZip&"'" &_
			",OwnerAddress='"&trim(request("WOwnerAddress"))&"'" &_
			" where BillNo='"&trim(request("WBillNo"))&"' and CarNo='"&trim(request("WCarNo"))&"'" 
	
		conn.execute strDciWUpdate
		ConnExecute strDciWUpdate,353
%>
<script language="JavaScript">
	alert("監理所回傳資料修改完成");
</script>
<%
end If

'修改DCI回傳資料(駕駛人)
if trim(request("kinds"))="DciW_Update2" then

	if trim(request("WDriverHomeAddress"))<>"" then
		strWDZip="select * from Zip where ZipName like '"&left(replace(trim(request("WDriverHomeAddress")),"臺","台"),5)&"%'"
		set rsWDZip=conn.execute(strWDZip)
		if not rsWDZip.eof then
			updWDriverZip=trim(rsWDZip("ZipID"))
		else
			if trim(request("WDriverHomeZip"))<>"" then
				updWDriverZip=trim(request("WDriverHomeZip"))
			else	
				updWDriverZip=""
			end if
		end if
		rsWDZip.close
		set rsWDZip=nothing
	end if
	'if sys_City="花蓮縣" Or sys_City="高雄縣" Or sys_City="高雄市" Or sys_City="彰化縣" Or sys_City="屏東縣" Or sys_City="台南市" Or sys_City="南投縣" or (sys_City="苗栗縣" And Trim(request("sys_BillTypeID"))="2") then
		strDciWUpdateA="update BillBaseDciReturn set " &_
			"Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverHomeZip='"&updWDriverZip&"'" &_
			",DriverHomeAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where CarNo='"&trim(request("WCarNo"))&"' and ExchangeTypeID='A'" 
		conn.execute strDciWUpdateA
	'end If
	'If sys_City="高雄市" or sys_City="花蓮縣" Or sys_City="南投縣" Or sys_City="台南市" Or sys_City="屏東縣" or sys_City="台中市" or sys_City="保二總隊三大隊一中隊" Then
		strUpdBillBase="Update BillBase set " &_
			"Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverZip='"&updWDriverZip&"'" &_
			",DriverAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where SN="&trim(request("BillSN"))
		conn.execute strUpdBillBase
	'End If 
		strDciWUpdate="update BillBaseDciReturn set " &_
			"Driver='"&trim(request("WDriver"))&"',DriverID='"&trim(request("WDriverID"))&"'" &_
			",DriverHomeZip='"&updWDriverZip&"'" &_
			",DriverHomeAddress='"&trim(request("WDriverHomeAddress"))&"'" &_
			" where BillNo='"&trim(request("WBillNo"))&"' and CarNo='"&trim(request("WCarNo"))&"'" 
	
		conn.execute strDciWUpdate
		ConnExecute strDciWUpdate,353
%>
<script language="JavaScript">
	alert("監理所回傳資料修改完成");
</script>
<%
end If

'修改建檔(車主)(查車)
if trim(request("kinds"))="CarQryUpdateKeyIn" then

	if trim(request("CarQryWOwnerAddress"))<>"" then
		strWOZip="select * from Zip where ZipName like '"&left(replace(trim(request("CarQryWOwnerAddress")),"臺","台"),5)&"%'"
		set rsWOZip=conn.execute(strWOZip)
		if not rsWOZip.eof then
			updWOwnerZip=trim(rsWOZip("ZipID"))
		else
			if trim(request("CarQryWOwnerZip"))<>"" then
				updWOwnerZip=trim(request("CarQryWOwnerZip"))
			else
				updWOwnerZip=""
			end if
		end if
		rsWOZip.close
		set rsWOZip=nothing
	end if

	strUpdBillBase="Update BillBase set " &_
		"Owner='"&trim(request("CarQryWOwner"))&"',OwnerID='"&trim(request("CarQryWOwnerID"))&"'" &_
		",OwnerZip='"&updWOwnerZip&"'" &_
		",OwnerAddress='"&trim(request("CarQryWOwnerAddress"))&"'" &_
		" where SN="&trim(request("BillSN"))
	conn.execute strUpdBillBase

	ConnExecute strUpdBillBase,353
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end If

'修改建檔資料(車主)
if trim(request("kinds"))="UpdateKeyInBillOwner" then
	updWOwnerZip=""
	if trim(request("BillOwnerAddress"))<>"" then
		strWOZip="select * from Zip where ZipName like '"&left(replace(trim(request("BillOwnerAddress")),"臺","台"),5)&"%'"
		set rsWOZip=conn.execute(strWOZip)
		if not rsWOZip.eof then
			updWOwnerZip=trim(rsWOZip("ZipID"))
		else
			if trim(request("BillOwnerZip"))<>"" then
				updWOwnerZip=trim(request("BillOwnerZip"))
			else
				updWOwnerZip=""
			end if
		end if
		rsWOZip.close
		set rsWOZip=nothing
	end if

	strUpdBillBase="Update BillBase set " &_
		"Owner='"&trim(request("BillOwner"))&"',OwnerID='"&trim(request("BillOwnerID"))&"'" &_
		",OwnerZip='"&updWOwnerZip&"'" &_
		",OwnerAddress='"&trim(request("BillOwnerAddress"))&"'" &_
		" where SN="&trim(request("BillSN"))
	conn.execute strUpdBillBase

	ConnExecute strUpdBillBase,353
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end If

'修改建檔駕駛人
if trim(request("kinds"))="UpdateKeyInBillDriver" then
	updWDriverZip=""
	if trim(request("BillDriverHomeAddress"))<>"" then
		strWDZip="select * from Zip where ZipName like '"&left(replace(trim(request("BillDriverHomeAddress")),"臺","台"),5)&"%'"
		set rsWDZip=conn.execute(strWDZip)
		if not rsWDZip.eof then
			updWDriverZip=trim(rsWDZip("ZipID"))
		else
			if trim(request("BillDriverHomeZip"))<>"" then
				updWDriverZip=trim(request("BillDriverHomeZip"))
			else	
				updWDriverZip=""
			end if
		end if
		rsWDZip.close
		set rsWDZip=nothing
	end if

		strUpdBillBase="Update BillBase set " &_
			"Driver='"&trim(request("BillDriver"))&"',DriverID='"&trim(request("BillDriverID2"))&"'" &_
			",DriverZip='"&updWDriverZip&"'" &_
			",DriverAddress='"&trim(request("BillDriverHomeAddress"))&"'" &_
			" where SN="&trim(request("BillSN"))
		conn.execute strUpdBillBase

		ConnExecute strUpdBillBase,353
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end If

'
'修改建檔駕駛人(入案)
if trim(request("kinds"))="CaseInUpdateKeyIn" then
	updWDriverZip=""
	if trim(request("CaseInDriverHomeAddress"))<>"" then
		strWDZip="select * from Zip where ZipName like '"&left(replace(trim(request("CaseInDriverHomeAddress")),"臺","台"),5)&"%'"
		set rsWDZip=conn.execute(strWDZip)
		if not rsWDZip.eof then
			updWDriverZip=trim(rsWDZip("ZipID"))
		else
			if trim(request("CaseInDriverHomeZip"))<>"" then
				updWDriverZip=trim(request("CaseInDriverHomeZip"))
			else	
				updWDriverZip=""
			end if
		end if
		rsWDZip.close
		set rsWDZip=nothing
	end if

		strUpdBillBase="Update BillBase set " &_
			"Driver='"&trim(request("CaseInDriver"))&"',DriverID='"&trim(request("CaseInDriverID"))&"'" &_
			",DriverZip='"&updWDriverZip&"'" &_
			",DriverAddress='"&trim(request("CaseInDriverHomeAddress"))&"'" &_
			" where SN="&trim(request("BillSN"))
		conn.execute strUpdBillBase

		ConnExecute strUpdBillBase,353
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end If

'上一筆
If Trim(request("kinds"))="Bill_PageUP" Then
	strUp="select Max(sn) as MaxSn from BillBase where RecordStateID=0 and Sn<"&Trim(request("BillSN"))&" and BillTypeID='1'" &_
		" and RecordMemberID="&Trim(session("User_ID"))
	'response.write strUp
	Set rsUp=conn.execute(strUp)
	If Not rsUp.eof Then
		 If trim(rsUp("MaxSn"))<>"" then
%>
<script language="JavaScript">
	location='BillKeyIn_Dci_Update.asp?BillSN=<%=Trim(rsUp("MaxSn"))%>';
</script>
<%		
		Else
%>
<script language="JavaScript">
	alert("已經是第一筆了!!");
</script>
<%		
		End If 
	End If 
	rsUp.close
	Set rsUp=Nothing 
End If 

'下一筆
If Trim(request("kinds"))="Bill_PageDown" Then
	strUp="select Min(sn) as MinSn from BillBase where RecordStateID=0 and Sn>"&Trim(request("BillSN"))&" and BillTypeID='1'" &_
		" and RecordMemberID="&Trim(session("User_ID"))
	'response.write strUp
	Set rsUp=conn.execute(strUp)
	If Not rsUp.eof Then
		If trim(rsUp("MinSn"))<>"" then
%>
<script language="JavaScript">
	location='BillKeyIn_Dci_Update.asp?BillSN=<%=Trim(rsUp("MinSn"))%>';
</script>
<%
		Else
%>
<script language="JavaScript">
	alert("已經是最後一筆了!!");
</script>
<%		
		End If 
	End If 
	rsUp.close
	Set rsUp=Nothing
End If 

'修改BillMailHistory
if trim(request("kinds"))="Mail_Update" then
	theDriverBirth=""
	if trim(request("StoreAndSendMailDate"))<>"" then
		theStoreAndSendMailDate=DateFormatChange(trim(request("StoreAndSendMailDate")))
	else 
		theStoreAndSendMailDate = "null"
	end if
	if trim(request("MailDate"))<>"" then
		theMailDate=DateFormatChange(trim(request("MailDate")))
	else 
		theMailDate = "null"
	end if
	if trim(request("OpenGovDate"))<>"" then
		theOpenGovDate=DateFormatChange(trim(request("OpenGovDate")))
	else 
		theOpenGovDate = "null"
	end if
	strMailUpdate="update BillMailHistory set " &_
		"MailDate="&theMailDate&",MailNumber='"&trim(request("MailNumber"))&"'" &_
		",StoreAndSendMailDate="&theStoreAndSendMailDate &_
		",StoreAndSendMailNumber='"&trim(request("StoreAndSendMailNumber"))&"'" &_
		",OpenGovDate="&theOpenGovDate &_
		" where BillSN="&trim(request("BillSN"))

		conn.execute strMailUpdate
		ConnExecute "送達記錄修改"&strMailUpdate,353
%>
<script language="JavaScript">
	alert("送達記錄修改完成");
</script>
<%
end If

If Trim(request("kinds"))="MemberStation_Update" Then
	strSUpd1="Update billbase set MemberStation='"&Trim(request("MemberStation"))&"' where SN="&trim(request("BillSN"))
	conn.execute strSUpd1

	strSUpd2="Update billbaseDcireturn set DciReturnStation='"&Trim(request("MemberStation"))&"' where BillNo='"&trim(request("SBillNo"))&"' and CarNo='"&trim(request("SCarNo"))&"'" 
	conn.execute strSUpd2

	strSUpd3="Update billbaseDcireturn set DciReturnStation='"&Trim(request("MemberStation"))&"' where ExchangeTypeID='A' and CarNo='"&trim(request("SCarNo"))&"'" 
	conn.execute strSUpd3

	ConnExecute "監理站修改"&strSUpd3,353

%>
<script language="JavaScript">
	alert("監理站修改完成");
</script>
<%
End If 

If Trim(request("kinds"))="CarType_Update" Then
	strUpd="Update Billbasedcireturn set DciReturnCarType='"&trim(request("DciCarType"))&"' where billno='"&trim(request("SBillNo"))&"' and carno='"&trim(request("SCarNo"))&"' and exchangetypeid='W'"
	conn.execute strUpd
	
	ConnExecute "詳細車種修改"&strUpd,353
%>
<script language="JavaScript">
	alert("詳細車種修改完成");
</script>
<%
End If 

If Trim(request("kinds"))="Update_ForFeit" Then	
	strForFeit=""
	If Trim(request("sys_ForFeit1"))<>"" Then
		strForFeit=" ForFeit1="&Trim(request("sys_ForFeit1"))
	End If 
	If Trim(request("sys_ForFeit2"))<>"" Then
		If strForFeit="" Then
			strForFeit=" ForFeit2="&Trim(request("sys_ForFeit2"))
		Else
			strForFeit=strForFeit&",ForFeit2="&Trim(request("sys_ForFeit2"))
		End If 
	End If 
	If Trim(request("sys_ForFeit3"))<>"" Then
		If strForFeit="" Then
			strForFeit=" ForFeit3="&Trim(request("sys_ForFeit3"))
		Else
			strForFeit=strForFeit&",ForFeit3="&Trim(request("sys_ForFeit3"))
		End If 
	End If 

	strUpd="Update Billbasedcireturn set "&strForFeit&" where billno='"&trim(request("SBillNo"))&"' and carno='"&trim(request("SCarNo"))&"' and exchangetypeid='W'"
	conn.execute strUpd
	
	'--------20220324 by jafe 高雄拖吊已結加入調整
	strUpd="Update Billbase set "&strForFeit&" where SN="&trim(request("BillSN"))
	conn.execute strUpd
	
	ConnExecute "修改罰款金額"&strUpd,353
%>
<script language="JavaScript">
	alert("罰款金額修改完成");
</script>
<%
End if
%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="right">
&nbsp; &nbsp; 登入者：<%=Session("Ch_Name") %>
<input type="button" value="離 開" onclick="window.close()"></div>
	<form name="myForm" method="post">
<%
if trim(request("theUpdVer"))<>"1" then
	strSqlbill="select * from BillBase where SN="&trim(request("BillSN"))
	'response.write strSqlbill
	set rsBill=conn.execute(strSqlbill)
	
	if trim(rsBill("BillTypeID"))<>"2" then
	'攔停
%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4"><strong>舉發單駕駛人資料修改</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;</td>
			</tr>
	<%if (session("ManagerPower"))="1" then%>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">單號</td>
				</td>
				<td>
					<input type="text" size="10" name="Billno1" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("BillNo"))
					else
						response.write trim(request("Billno1"))
					end if					
					%>" onblur="CheckBillNoExist()" maxlength="9">

					<input type="hidden" name="OldBillNo" value="<%response.write trim(rsBill("BillNo"))%>">
				</td>
				<td bgcolor="#EBE5FF" width="15%">車牌號碼</td>
				<td>
					<input type="text" size="10" name="NewCarNo" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("CarNo"))
					else
						response.write trim(request("NewCarNo"))
					end if					
					%>" onblur="getVIPCar()">

					<input type="hidden" name="OldCarNo" value="<%response.write trim(rsBill("CarNo"))%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">駕駛人生日</td>
				<td width="35%">
					<input type="text" name="BillDriverBirth" size="10" maxlength="7" value="<%
					if trim(request("kinds"))="" then
						if trim(rsBill("DriverBirth"))<>"" and not isnull(rsBill("DriverBirth")) then
							response.write ginitdt(trim(rsBill("DriverBirth")))
						end if
					else
						response.write trim(request("BillDriverBirth"))
					end if
					%>">
				</td>
				<td bgcolor="#EBE5FF" width="15%">駕駛人證號</td>
				<td width="35%">
					<input type="text" size="10" name="BillDriverID" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("DriverID"))
					else
						response.write trim(request("BillDriverID"))
					end if					
					%>">
				</td>
			</tr>
	<%else%>
					<input type="hidden" size="10" name="Billno1" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("BillNo"))
					else
						response.write trim(request("Billno1"))
					end if					
					%>" onblur="CheckBillNoExist()" maxlength="9">
					<input type="hidden" name="OldBillNo" value="<%response.write trim(rsBill("BillNo"))%>">
					<input type="hidden" size="10" name="NewCarNo" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("CarNo"))
					else
						response.write trim(request("NewCarNo"))
					end if					
					%>" onblur="getVIPCar()">
					<input type="hidden" name="OldCarNo" value="<%response.write trim(rsBill("CarNo"))%>">
					<input type="hidden" name="BillDriverBirth" size="10" maxlength="7" value="<%
					if trim(request("kinds"))="" then
						if trim(rsBill("DriverBirth"))<>"" and not isnull(rsBill("DriverBirth")) then
							response.write ginitdt(trim(rsBill("DriverBirth")))
						end if
					else
						response.write trim(request("BillDriverBirth"))
					end if
					%>">
					<input type="hidden" size="10" name="BillDriverID" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("DriverID"))
					else
						response.write trim(request("BillDriverID"))
					end if					
					%>">

	<%end if%>
			<tr>
				<td bgcolor="#EBE5FF"><font color="red">備註</font></td>
				<td colspan="3">
					<input type="text" name="Note" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("Note"))
					else
						response.write trim(request("Note"))
					end if					
					%>" size="50">
				</td>
			</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="4">
					<input type="button" value="儲 存" onclick="InsertBillVase();" class="btn1">
					<!-- 違規人性別 -->
					<input type="hidden" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("DriverSex"))
					else
						response.write trim(request("BillDriverSex"))
					end if					
					%>" name="BillDriverSex">
				</td>
			</tr>
		</table>	
		<br>
<%
	else
	'逕舉
%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4"><strong>舉發單資料修改</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;</td>
			</tr>
	<%if (session("ManagerPower"))="1" then%>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">車牌號碼</td>
				<td colspan="3">
					<input type="text" size="10" name="NewCarNo" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("CarNo"))
					else
						response.write trim(request("NewCarNo"))
					end if					
					%>" onblur="getVIPCar()">

					<input type="hidden" name="OldCarNo" value="<%response.write trim(rsBill("CarNo"))%>">
					<input type="hidden" size="10" name="Billno1" value="<%
						response.write trim(rsBill("BillNo"))
					%>">

					<input type="hidden" name="OldBillNo" value="<%response.write trim(rsBill("BillNo"))%>">
				</td>
				<!-- <td bgcolor="#EBE5FF" width="15%">駕駛人證號</td>
				<td width="35%">
					
				</td> -->
			</tr>
	<%else%>
			<tr>
				<td colspan="3">
					<input type="hidden" size="10" name="NewCarNo" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("CarNo"))
					else
						response.write trim(request("NewCarNo"))
					end if					
					%>" onblur="getVIPCar()">

					<input type="hidden" name="OldCarNo" value="<%response.write trim(rsBill("CarNo"))%>">
					<input type="hidden" size="10" name="Billno1" value="<%
						response.write trim(rsBill("BillNo"))
					%>">

					<input type="hidden" name="OldBillNo" value="<%response.write trim(rsBill("BillNo"))%>">
				<!-- <td bgcolor="#EBE5FF" width="15%">駕駛人證號</td>
				<td width="35%">
					
				</td> -->
			</tr>
	<%end if%>
			<tr>
				<td bgcolor="#EBE5FF"><font color="red">備註</font></td>
				<td colspan="3">
					<input type="text" name="Note" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("Note"))
					else
						response.write trim(request("Note"))
					end if					
					%>" size="50">
				</td>
			</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="4">
					<input type="button" value="儲 存" onclick="InsertBillVase();" class="btn1">
					<!-- 違規人性別 -->
					<input type="hidden" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("DriverSex"))
					else
						response.write trim(request("BillDriverSex"))
					end if					
					%>" name="BillDriverSex">
					<!-- 駕駛人生日 -->
					<input type="hidden" name="BillDriverBirth" size="10" maxlength="7" value="<%
					if trim(request("kinds"))="" then
						if trim(rsBill("DriverBirth"))<>"" and not isnull(rsBill("DriverBirth")) then
							response.write ginitdt(trim(rsBill("DriverBirth")))
						end if
					else
						response.write trim(request("BillDriverBirth"))
					end if
					%>">
					<!-- 駕駛人證號 -->
					<input type="hidden" size="10" name="BillDriverID" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsBill("DriverID"))
					else
						response.write trim(request("BillDriverID"))
					end if					
					%>">
				</td>
			</tr>
		</table>	
		<br>
<%
	end If

	if sys_City="苗栗縣" then %>
	<table width='985' border='1' align="center" cellpadding="1">
		<tr bgcolor="#FFCCCC">
			<td colspan="4"><strong>舉發單建檔地址修改</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;</td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="15%">車主姓名</td>
			<td width="35%">
				<input type="text" name="BillOwner" size="20" maxlength="50" value="<%
					if trim(rsBill("Owner"))<>"" and not isnull(rsBill("Owner")) then
						response.write trim(rsBill("Owner"))
					end if
				%>" style=ime-mode:active>
			</td>
			<td bgcolor="#EBE5FF" width="15%">車主身分證號</td>
			<td width="35%">
				<input type="text" size="10" name="BillOwnerID" value="<%
					response.write trim(rsBill("OwnerID"))
				%>" style=ime-mode:disabled onkeyup="this.value=this.value.toUpperCase();">
			</td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="15%">車主地址</td>
			<td colspan="3">
				郵遞區號
				<input type="text" name="BillOwnerZip" size="5" value="<%
					if trim(rsBill("OwnerZip"))<>"" and not isnull(rsBill("OwnerZip")) then
						response.write trim(rsBill("OwnerZip"))
					end if
				%>"  style=ime-mode:disabled>
				地址
				<input type="text" name="BillOwnerAddress" size="60" value="<%
					if trim(rsBill("OwnerAddress"))<>"" and not isnull(rsBill("OwnerAddress")) then
						response.write trim(rsBill("OwnerAddress"))
					end if

				%>" style=ime-mode:active>
				<input type="button" value="儲存車主(建檔)" onclick="UpdateKeyInBillOwner();" class="btn1">
			</td>
		</tr>
<%
	strDciW2="select b.Billno,b.CarNo,b.Owner,b.OwnerID,b.OwnerAddress,b.OwnerZip,b.Driver" &_
	",b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,a.ExchangeTypeID,b.DcireturnCarType" &_
	" from DciLog a,BillBaseDciReturn b" &_
	" where a.BillSn="&trim(request("BillSN"))&_ 
	" and (b.BillNo is null and a.CarNo=b.CarNo and a.ExchangeTypeID='A')" &_
	" and a.ExchangeTypeID=b.ExchangeTypeID and a.DciReturnStatusID=b.Status" &_
	" order by a.ExchangeTypeID Desc"
	Set rsDciW2=conn.execute(strDciW2)
	If Not rsDciW2.eof then
%>
		<tr>
			<td bgcolor="#CCFFFF" width="15%">(查車)車主姓名</td>
			<td width="35%">
				<%
					if trim(rsDciW2("Owner"))<>"" and not isnull(rsDciW2("Owner")) then
						response.write trim(rsDciW2("Owner"))
					end if
				%>
				<input type="hidden" name="CarQryWOwner" size="20" maxlength="50" value="<%
					if trim(rsDciW2("Owner"))<>"" and not isnull(rsDciW2("Owner")) then
						response.write trim(rsDciW2("Owner"))
					end if
				%>">
			</td>
			<td bgcolor="#CCFFFF" width="15%">(查車)身分證號</td>
			<td width="35%">
				<%
					response.write trim(rsDciW2("OwnerID"))
				%>
				<input type="hidden" size="10" name="CarQryWOwnerID" value="<%
					response.write trim(rsDciW2("OwnerID"))
				%>">
			</td>
		</tr>
		<tr>
			<td bgcolor="#CCFFFF" width="15%">(查車)車主地址</td>
			<td colspan="3">
				<%
					if trim(rsDciW2("OwnerZip"))<>"" and not isnull(rsDciW2("OwnerZip")) then
						response.write trim(rsDciW2("OwnerZip"))
					end if
				%>
				<input type="hidden" name="CarQryWOwnerZip" size="5" value="<%
					if trim(rsDciW2("OwnerZip"))<>"" and not isnull(rsDciW2("OwnerZip")) then
						response.write trim(rsDciW2("OwnerZip"))
					end if
				%>">
				<%
					if trim(rsDciW2("OwnerAddress"))<>"" and not isnull(rsDciW2("OwnerAddress")) then
						response.write trim(rsDciW2("OwnerAddress"))
					end if
				%>
				<input type="hidden" name="CarQryWOwnerAddress" size="60" value="<%
					if trim(rsDciW2("OwnerAddress"))<>"" and not isnull(rsDciW2("OwnerAddress")) then
						response.write trim(rsDciW2("OwnerAddress"))
					end if
				%>">
				<input type="button" value="用查車資料儲存車主(建檔)" onclick="CarQryUpdateKeyIn();" class="btn1">
			</td>
		</tr>
<%
	End If 
%>
		<tr>
			<td bgcolor="#EBE5FF" width="15%">駕駛人姓名</td>
			<td width="35%">
				<input type="text" name="BillDriver" size="20" maxlength="50" value="<%
					if trim(rsBill("Driver"))<>"" and not isnull(rsBill("Driver")) then
						response.write trim(rsBill("Driver"))
					end if
				%>" style=ime-mode:active >
			</td>
			<td bgcolor="#EBE5FF" width="15%">駕駛人身分證號</td>
			<td width="35%">
				<input type="text" size="10" name="BillDriverID2" value="<%
					response.write trim(rsBill("DriverID"))
			
				%>" style=ime-mode:disabled onkeyup="this.value=this.value.toUpperCase();">
			</td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="15%">駕駛人地址</td>
			<td colspan="3">
				郵遞區號
				<input type="text" name="BillDriverHomeZip" size="5" value="<%

					sysBillTypeID=trim(rsBill("BillTypeID"))
					sysCarNo=trim(rsBill("CarNo"))


				sysDriverHomeAddress=""
				if trim(rsBill("DriverZip"))<>"" and not isnull(rsBill("DriverZip")) then
					response.write trim(rsBill("DriverZip"))
				end If
				
				%>" style=ime-mode:disabled>
				地址
				<input type="text" name="BillDriverHomeAddress" size="60" value="<%
				if trim(rsBill("DriverAddress"))<>"" and not isnull(rsBill("DriverAddress")) then
					response.write trim(rsBill("DriverAddress"))
				end if
				
				%>" style=ime-mode:active>
				<input type="button" value="儲存駕駛人(建檔)" onclick="UpdateKeyInBillDriver();" class="btn1">
			</td>
		</tr>
<%
	strDciW2="select a.BillTypeID,b.Billno,b.CarNo,b.Owner,b.OwnerID,b.OwnerAddress,b.OwnerZip,b.Driver" &_
		",b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,b.DcireturnCarType" &_
		" from DciLog a,BillBaseDciReturn b" &_
		" where a.BillSn="&trim(request("BillSN"))&" and a.BillNo=b.BillNo and a.CarNo=b.CarNo" &_
		" and a.ExchangeTypeID=b.ExchangeTypeID and a.DciReturnStatusID=b.Status" &_
		" and a.ExchangeTypeID='W'"
	Set rsDciW2=conn.execute(strDciW2)
	If Not rsDciW2.eof then
%>
		<tr>
			<td bgcolor="#CCFFFF" width="15%">(入案)駕駛人姓名</td>
			<td width="35%">
				<%
					if trim(rsDciW2("Driver"))<>"" and not isnull(rsDciW2("Driver")) then
						response.write trim(rsDciW2("Driver"))
					end if
				%>
				<input type="hidden" name="CaseInDriver" size="20" maxlength="14" value="<%
					if trim(rsDciW2("Driver"))<>"" and not isnull(rsDciW2("Driver")) then
						response.write trim(rsDciW2("Driver"))
					end if
				%>">
			</td>
			<td bgcolor="#CCFFFF" width="15%">(入案)身分證號</td>
			<td width="35%">
				<%
					response.write trim(rsDciW2("DriverID"))
				%>
				<input type="hidden" size="10" name="CaseInDriverID" value="<%
					response.write trim(rsDciW2("DriverID"))
				%>">
			</td>
		</tr>
		<tr>
			<td bgcolor="#CCFFFF" width="15%">(入案)駕駛人地址</td>
			<td colspan="3">
				<%
					if trim(rsDciW2("DriverHomeZip"))<>"" and not isnull(rsDciW2("DriverHomeZip")) then
						response.write trim(rsDciW2("DriverHomeZip"))
					end if
				%>
				<input type="hidden" name="CaseInDriverHomeZip" size="5" value="<%
					if trim(rsDciW2("DriverHomeZip"))<>"" and not isnull(rsDciW2("DriverHomeZip")) then
						response.write trim(rsDciW2("DriverHomeZip"))
					end if
				%>">
				<%
					if trim(rsDciW2("DriverHomeAddress"))<>"" and not isnull(rsDciW2("DriverHomeAddress")) then
						response.write trim(rsDciW2("DriverHomeAddress"))
					end if
				%>
				<input type="hidden" name="CaseInDriverHomeAddress" size="60" value="<%
					if trim(rsDciW2("DriverHomeAddress"))<>"" and not isnull(rsDciW2("DriverHomeAddress")) then
						response.write trim(rsDciW2("DriverHomeAddress"))
					end if
				%>">
				<input type="button" value="用入案資料儲存駕駛人(建檔)" onclick="CaseInUpdateKeyIn();" class="btn1">
			</td>
		</tr>
<%		If trim(rsBill("BillTypeID"))="1" then%>
		<tr>
			<td colspan="4" align="center" bgcolor="#FFCCCC">
				<input type="button" value="上一筆" name="btn_PageUP" style="background-color:#B9FBBE;" onclick="Bill_PageUP();">&nbsp; &nbsp; &nbsp; 
				<input type="button" value="下一筆" name="btn_PageDown" style="background-color:#B9FBBE;" onclick="Bill_PageDown();">
			</td>
		</tr>
<%		End If %>
<%
	End If 
%>
	</table>	
	<br>
	<%
	End If 
	rsBill.close
	set rsBill=nothing
end if
%>

<%'監理所回傳資料修改
if (session("ManagerPower"))="1" or trim(request("theUpdVer"))="1" Or sys_City="苗栗縣" Then
	DciCarTypeID=""
	ForFeit1=0
	ForFeit2=0
	ForFeit3=0
	Rule1=""
	Rule2=""
	Rule3=""
	if sys_City="花蓮縣" then
		strDciW="select a.BillTypeID,b.Billno,b.CarNo,b.Owner,b.OwnerID,b.OwnerAddress,b.OwnerZip,b.Driver" &_
		",b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,a.ExchangeTypeID,b.DcireturnCarType" &_
		",b.ForFeit1,b.ForFeit2,b.ForFeit3,b.Rule1,b.Rule2,b.Rule3" &_
		" from DciLog a,BillBaseDciReturn b" &_
		" where a.BillSn="&trim(request("BillSN"))&" and ((a.BillNo=b.BillNo and a.CarNo=b.CarNo" &_
		" and a.ExchangeTypeID='W') " &_
		" or (b.BillNo is null and a.CarNo=b.CarNo and a.ExchangeTypeID='A'))" &_
		" and a.ExchangeTypeID=b.ExchangeTypeID and a.DciReturnStatusID=b.Status" &_
		" order by a.ExchangeTypeID Desc"
	else
		strDciW="select a.BillTypeID,b.Billno,b.CarNo,b.Owner,b.OwnerID,b.OwnerAddress,b.OwnerZip,b.Driver" &_
		",b.DriverID,b.DriverHomeZip,b.DriverHomeAddress,b.DcireturnCarType" &_
		",b.ForFeit1,b.ForFeit2,b.ForFeit3,b.Rule1,b.Rule2,b.Rule3" &_
		" from DciLog a,BillBaseDciReturn b" &_
		" where a.BillSn="&trim(request("BillSN"))&" and a.BillNo=b.BillNo and a.CarNo=b.CarNo" &_
		" and a.ExchangeTypeID=b.ExchangeTypeID and a.DciReturnStatusID=b.Status" &_
		" and a.ExchangeTypeID='W'"
	end if
	set rsDciW=conn.execute(strDciW)
	if not rsDciW.eof Then
		DciCarTypeID=Trim(rsDciW("DcireturnCarType"))
		ForFeit1=Trim(rsDciW("ForFeit1"))
		ForFeit2=Trim(rsDciW("ForFeit2"))
		ForFeit3=Trim(rsDciW("ForFeit3"))
		Rule1=Trim(rsDciW("Rule1"))
		Rule2=Trim(rsDciW("Rule2"))
		Rule3=Trim(rsDciW("Rule3"))
%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4"><strong>監理所回傳資料修改</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="button" name="b1" value="路名查詢" onclick='window.open("../AddressQry.asp","AddressQry","left=500,top=150,location=0,width=600,height=400,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 80px; height:26px;">
				</a>
				<input type="hidden" name="sys_BillTypeID" value="<%=trim(rsDciW("BillTypeID"))%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">車主姓名</td>
				<td width="35%">
					<input type="text" name="WOwner" size="20" maxlength="50" value="<%
						if trim(rsDciW("Owner"))<>"" and not isnull(rsDciW("Owner")) then
							response.write trim(rsDciW("Owner"))
						end if
					%>" style=ime-mode:active>
				</td>
				<td bgcolor="#EBE5FF" width="15%">車主身分證號</td>
				<td width="35%">
					<input type="text" size="10" name="WOwnerID" value="<%
						response.write trim(rsDciW("OwnerID"))
					%>" style=ime-mode:disabled onkeyup="this.value=this.value.toUpperCase();">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">車主地址</td>
				<td colspan="3">
					郵遞區號
					<input type="text" name="WOwnerZip" size="5" value="<%
						if trim(rsDciW("OwnerZip"))<>"" and not isnull(rsDciW("OwnerZip")) then
							response.write trim(rsDciW("OwnerZip"))
						end if
					%>" style=ime-mode:disabled>
					地址
					<input type="text" name="WOwnerAddress" size="60" value="<%
						if trim(rsDciW("OwnerAddress"))<>"" and not isnull(rsDciW("OwnerAddress")) then
							response.write trim(rsDciW("OwnerAddress"))
						end if
	
					%>" style=ime-mode:active>
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">駕駛人姓名</td>
				<td width="35%">
					<input type="text" name="WDriver" size="20" maxlength="50" value="<%
					if trim(request("kinds"))="" then
						if trim(rsDciW("Driver"))<>"" and not isnull(rsDciW("Driver")) then
							response.write trim(rsDciW("Driver"))
						end if
					else
						response.write trim(request("WDriver"))
					end if
					%>" style=ime-mode:active>
				</td>
				<td bgcolor="#EBE5FF" width="15%">駕駛人身分證號</td>
				<td width="35%">
					<input type="text" size="10" name="WDriverID" value="<%
					if trim(request("kinds"))="" then
						response.write trim(rsDciW("DriverID"))
					else
						response.write trim(request("WDriverID"))
					end if					
					%>" style=ime-mode:disabled onkeyup="this.value=this.value.toUpperCase();">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="15%">駕駛人地址</td>
				<td colspan="3">
					郵遞區號
					<input type="text" name="WDriverHomeZip" size="5" value="<%
					strSqlbill2="select CarNo,BillNo,BillTypeID,DriverBirth,DriverID,DriverSex,RecordMemberID,Note from BillBase where SN="&trim(request("BillSN"))
					Set rsbill2=conn.execute(strSqlbill2)
					If Not rsbill2.eof Then
						sysBillTypeID=trim(rsbill2("BillTypeID"))
						sysCarNo=trim(rsbill2("CarNo"))
					End If 
					rsbill2.close
					Set rsbill2=nothing

					sysDriverHomeAddress=""
					if sysBillTypeID<>"2" then
						if trim(rsDciW("DriverHomeZip"))<>"" and not isnull(rsDciW("DriverHomeZip")) then
							response.write trim(rsDciW("DriverHomeZip"))
						end If
					Else
						if trim(rsDciW("DriverHomeZip"))<>"" and not isnull(rsDciW("DriverHomeZip")) then
							response.write trim(rsDciW("DriverHomeZip"))
'						Else
'							strDciA="select * from BillBaseDciReturn where ExchangeTypeID='A' and Status='S'" &_
'								" and CarNo='"&sysCarNo&"'"
'							Set rsDciA=conn.execute(strDciA)
'							If Not rsDciA.eof Then
'								if trim(rsDciA("DriverHomeZip"))<>"" and not isnull(rsDciA("DriverHomeZip")) then
'									response.write trim(rsDciA("DriverHomeZip"))
'								end If
'								sysDriverHomeAddress=trim(rsDciA("DriverHomeAddress"))
'							End if
'							rsDciA.close
'							Set rsDciA=nothing
						end If
					End If
					
					%>" style=ime-mode:disabled>
					地址
					<input type="text" name="WDriverHomeAddress" size="60" value="<%
					if sysBillTypeID<>"2" then
						if trim(rsDciW("DriverHomeAddress"))<>"" and not isnull(rsDciW("DriverHomeAddress")) then
							response.write trim(rsDciW("DriverHomeAddress"))
						end if
					Else
						if trim(rsDciW("DriverHomeAddress"))<>"" and not isnull(rsDciW("DriverHomeAddress")) then
							response.write trim(rsDciW("DriverHomeAddress"))
'						Else
'							response.write sysDriverHomeAddress
						end if
					End if
						
					%>" style=ime-mode:active>
				</td>
			</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="4">
					<input type="button" value="儲存車主" onclick="UpdateDciW1();" <%

					%> class="btn1">
					<input type="button" value="儲存駕駛人" onclick="UpdateDciW2();" <%

					%> class="btn1">
					<input type="button" value="全部儲存" onclick="UpdateDciW();" <%

					%> class="btn1">
					<input type="hidden" name="WBillno" value="<%=trim(rsDciW("Billno"))%>">
					<input type="hidden" name="WCarNo" value="<%=trim(rsDciW("CarNo"))%>">
				</td>
			</tr>
		</table>
		<br>
<%	end if
	rsDciW.close
	set rsDciW=nothing
end if%>

<%'監理所送達資料修改
if (session("ManagerPower"))="1" or trim(request("theUpdVer"))="1" then
	'BillMailHisory
	strHis="select * from BillMailHistory where BillSn="&trim(request("BillSN"))
	set rsHis=conn.execute(strHis)
	if not rsHis.eof then
%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4"><strong>送達紀錄修改</strong></td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF">郵寄日期</td>
				<td>
					<input type="text" name="MailDate" size="10" maxlength="7" value="<%
						if trim(rsHis("MailDate"))<>"" and not isnull(rsHis("MailDate")) then
							response.write ginitdt(trim(rsHis("MailDate")))
						end if
					%>">
				</td>
				<td bgcolor="#EBE5FF">郵寄序號</td>
				<td>
					<input type="text" size="10" name="MailNumber" value="<%
					if trim(rsHis("MailNumber"))<>"" and not isnull(rsHis("MailNumber")) then
						response.write trim(rsHis("MailNumber"))
					end if
					%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="20%">寄存送達生效(完成)日</td>
				<td width="30%">
					<input type="text" name="StoreAndSendMailDate" size="10" maxlength="7" value="<%
						if trim(rsHis("StoreAndSendMailDate"))<>"" and not isnull(rsHis("StoreAndSendMailDate")) then
							response.write ginitdt(trim(rsHis("StoreAndSendMailDate")))
						end if
					%>">
				</td>
				<td bgcolor="#EBE5FF" width="20%">寄存送達掛號碼</td>
				<td width="30%">
					<input type="text" size="10" name="StoreAndSendMailNumber" value="<%
					if trim(rsHis("StoreAndSendMailNumber"))<>"" and not isnull(rsHis("StoreAndSendMailNumber")) then
						response.write trim(rsHis("StoreAndSendMailNumber"))
					end if					
					%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="20%">公示送達生效日</td>
				<td colspan="3">
					<input type="text" name="OpenGovDate" size="10" maxlength="7" value="<%
						if trim(rsHis("OpenGovDate"))<>"" and not isnull(rsHis("OpenGovDate")) then
							response.write ginitdt(trim(rsHis("OpenGovDate")))
						end if
					%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="4">
					<input type="button" value="儲 存" onclick="UpdateMail();" <%

					%> class="btn1">
				</td>
			</tr>
		</table>
<%	end if
	rsHis.close
	set rsHis=nothing
end If
If Trim(Session("Credit_ID"))="A000000000" or trim(request("theUpdVer"))="1" Then
%>
	<br>
	<table width='985' border='1' align="center" cellpadding="1">
		<tr bgcolor="#1BF5FF">
			<td colspan="4"><strong>詳細車種修改</strong></td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="20%">詳細車種</td>
			<td colspan="3">
				<select name="DciCarType">
<%
	strT="select * from DciCode where TypeID=5 order by ID"
	Set rsT=conn.execute(strT)
	If Not rsT.Bof Then rsT.MoveFirst 
	While Not rsT.Eof
%>
					<option value="<%=Trim(rsT("ID"))%>" <%
					If DciCarTypeID=Trim(rsT("ID")) Then
						response.write "selected"
					End If 
					%>><%=Trim(rsT("Content"))%></option>
<%	rsT.MoveNext
	Wend
	rsT.close
	set rsT=nothing
%>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="#1BF5FF" align="center" colspan="4">
				<input type="button" value="儲 存" onclick="UpdateCarType();" class="btn1">
			</td>
		</tr>
	</table>
	<br>
	<table width='985' border='1' align="center" cellpadding="1">
		<tr bgcolor="#1BF5FF">
			<td colspan="4"><strong>監理站修改</strong></td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="20%">監理站</td>
			<td colspan="3">
<%	MemberStationTmp=""
	strS1="select CarNo,BillNo,BillTypeID,MemberStation from BillBase where SN="&trim(request("BillSN"))
	Set rsS1=conn.execute(strS1)
	If Not rsS1.eof Then
		If Trim(rsS1("BillTypeID"))="1" Then
			MemberStationTmp=Trim(rsS1("MemberStation"))
		Else
			strS2="select * from BillBaseDcireturn where BillNo='"&Trim(rsS1("BillNo"))&"' and CarNo='"&Trim(rsS1("CarNo"))&"' and exchangetypeid='W'"
			Set rsS2=conn.execute(strS2)
			If Not rsS2.eof Then
				MemberStationTmp=Trim(rsS2("DciReturnStation"))
			End If
			rsS2.close
			Set rsS2=Nothing 
		End If
%>
				<input type="hidden" name="SBillno" value="<%=trim(rsS1("Billno"))%>">
				<input type="hidden" name="SCarNo" value="<%=trim(rsS1("CarNo"))%>">
<%
	End If
	rsS1.close
	Set rsS1=Nothing 
%>
				<select name="MemberStation">
					<option value="">查無監理站</option>
<%
	strS3="select a.DciStationID,a.DciStationName,a.StationAddress,a.StationTel from Station a," &_
			"(select distinct(StationID) from Station) b where a.DciStationID=b.StationID order by a.StationID"
	Set rsS3=conn.execute(strS3)
	If Not rsS3.eof Then rsS3.MoveFirst
		While Not rsS3.Eof
%>
					<option value="<%=Trim(rsS3("DciStationID"))%>" <%
					If Trim(rsS3("DciStationID"))=MemberStationTmp Then
						response.write "selected"
					End If 
					%>><%=Trim(rsS3("DciStationID"))&" "&trim(rsS3("DciStationName"))%></option>
<%
		rsS3.MoveNext
		Wend
	rsS3.close
	Set rsS3=Nothing 
%>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="#1BF5FF" align="center" colspan="4">
				<input type="button" value="儲 存" onclick="UpdateMemStation();" class="btn1">
			</td>
		</tr>
	</table>
	<br>
	<table width='985' border='1' align="center" cellpadding="1">
		<tr bgcolor="#1BF5FF">
			<td colspan="4"><strong>罰款金額修改</strong></td>
		</tr>
<%If Trim(Rule1)<>"" And Trim(Rule1)<>"0" then%>
		<tr>
			<td bgcolor="#EBE5FF" width="20%">法條一 </td>
			<td colspan="3">
				( <%=Trim(Rule1)%> ) <input type="text" value="<%=Trim(ForFeit1)%>" name="sys_ForFeit1">
			</td>
		</tr>
<%End If %>
<%If Trim(Rule2)<>"" And Trim(Rule2)<>"0" then%>
		<tr>
			<td bgcolor="#EBE5FF" width="20%">法條二 </td>
			<td colspan="3">
				( <%=Trim(Rule2)%> ) <input type="text" value="<%=Trim(ForFeit2)%>" name="sys_ForFeit2">
			</td>
		</tr>
<%End If %>
<%If Trim(Rule3)<>"" And Trim(Rule3)<>"0" then%>
		<tr>
			<td bgcolor="#EBE5FF" width="20%">法條三 </td>
			<td colspan="3">
				( <%=Trim(Rule3)%> ) <input type="text" value="<%=Trim(ForFeit3)%>" name="sys_ForFeit3">
			</td>
		</tr>
<%End If %>
		<tr>
			<td bgcolor="#1BF5FF" align="center" colspan="4">
				<input type="button" value="儲 存" onclick="UpdateForFeit();" class="btn1">
			</td>
		</tr>
	</table>
<%
End if
%>
		<input type="hidden" value="" name="kinds">
		<input type="hidden" value="<%=trim(request("BillSN"))%>" name="BillSN">

	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
var TDMemErrorLog1=0;
//修改告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	if (myForm.BillDriverBirth.value!=""){
		if (!dateCheck( myForm.BillDriverBirth.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：駕駛人生日輸入錯誤。";
		}
	}
	if (error==0){
		myForm.kinds.value="DB_insert";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改DCI入案
function UpdateDciW(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="DciW_Update";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
<%if sys_City="苗栗縣" then%>
//修改建檔車主(用查車資料)
function CarQryUpdateKeyIn(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="CarQryUpdateKeyIn";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改建檔車主
function UpdateKeyInBillOwner(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="UpdateKeyInBillOwner";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改建檔駕駛
function UpdateKeyInBillDriver(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="UpdateKeyInBillDriver";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改建檔駕駛(入案)
function CaseInUpdateKeyIn(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="CaseInUpdateKeyIn";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
function Bill_PageUP(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="Bill_PageUP";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
function Bill_PageDown(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="Bill_PageDown";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
<%end if%>
//修改DCI入案
function UpdateDciW1(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="DciW_Update1";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改DCI入案
function UpdateDciW2(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="DciW_Update2";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改DCI寄存
function UpdateDciNF(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="DciNF_Update";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改DCI公示
function UpdateDciND(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="DciND_Update";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//修改BILLMAILHISTORY
function UpdateMail(){
	var error=0;
	var errorString="";
	if (myForm.MailDate.value!=""){
		if (!dateCheck( myForm.MailDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規通知單郵寄日期輸入錯誤。";
		}
	}
	if (myForm.StoreAndSendMailDate.value!=""){
		if (!dateCheck( myForm.StoreAndSendMailDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：寄存送達投郵日期輸入錯誤。";
		}
	}
	if (myForm.OpenGovDate.value!=""){
		if (!dateCheck( myForm.OpenGovDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：公示紀錄日期輸入錯誤。";
		}
	}
	if (error==0){
		myForm.kinds.value="Mail_Update";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
function FuncChkPID(){
	myForm.DriverPID.value=myForm.DriverPID.value.toUpperCase();
	if (myForm.DriverPID.value.length == 10){
		if (!check_tw_id(myForm.DriverPID.value)){
			alert("身分證輸入錯誤！");
		}else{
			if (myForm.DriverPID.value.substr(1,1)=="1"){
				document.myForm.DriverSex.value="1";
			}else{
				document.myForm.DriverSex.value="2";
			}
		}
	}
}
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value != ""){
	//alert(myForm.Billno1.value)
		if (myForm.Billno1.value.length=="9" ){
			if (myForm.Billno1.value!=myForm.OldBillNo.value) {
				runServerScript("getCheckBillNoExistforUpdate.asp?BillNo="+BillNum);
				
			}
		}else{
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}
	}
}


//是否為特殊用車&檢查是否有同車號在同一天建檔
function getVIPCar(){
	myForm.NewCarNo.value=myForm.NewCarNo.value.toUpperCase();
	myForm.NewCarNo.value=myForm.NewCarNo.value.replace(" ", "");
	var CarNum=myForm.NewCarNo.value;
//	CarType=chkCarNoFormat(myForm.NewCarNo.value);
//	if (CarType==0){
//		alert("車牌格式錯誤");
//		myForm.NewCarNo.select();
//	}
}

<%If Trim(Session("Credit_ID"))="A000000000" or trim(request("theUpdVer"))="1" Then%>
function UpdateMemStation(){
	if (myForm.MemberStation.value==""){
		alert("請選擇監理站!!");
	}else{
		myForm.kinds.value="MemberStation_Update";
		myForm.submit();
	}
}

function UpdateCarType(){
	if (myForm.DciCarType.value==""){
		alert("請選擇詳細車種!!");
	}else{
		myForm.kinds.value="CarType_Update";
		myForm.submit();
	}
}
//修改罰款金額
function UpdateForFeit(){
	var error=0;
	var errorString="";
	if (error==0){
		myForm.kinds.value="Update_ForFeit";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
<%end if%>
</script>
</html>
