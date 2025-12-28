<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
On Error Resume Next
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
If sys_City="宜蘭縣" Then

	fname=year(now)&fMnoth&fDay&"_行政執行處移送電子清冊.odf"

else

	fname=year(now)&fMnoth&fDay&"_行政執行處移送電子清冊.xls"
End if 

Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950"

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If


strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
rsUInfo.close
set rsUInfo=nothing

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	DB_UnitID=trim(unit("UnitID"))
	if not isnull(unit("UnitName")) and trim(unit("UnitName"))<>"" then
		DB_UnitName=replace(replace(trim(unit("UnitName")),"交通組",""),"台","臺")
	end if 
	DB_Tel=trim(unit("Tel"))
	theSubUnitSecBossName=trim(unit("SecondManagerName"))
	theBigUnitBossName=trim(unit("ManageMemberName"))
	thePasserSendBankAccountName=trim(unit("PasserSendBankAccountName"))
	thePasserVATnumber=trim(unit("VATNUMBER"))
	thePasserSendBankAccount=trim(unit("PasserSendBankAccount"))
	thePasserSendBankName=trim(unit("PasserSendBankName"))
	theBankName=trim(unit("BankName"))
	theBankAccount=trim(unit("BankAccount"))
end if
unit.close

If ifnull(thePasserSendBankAccount) Then
	thePasserSendBankAccount=trim(theBankAccount)
	thePasserSendBankName=trim(theBankName)
End if

If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"


	strSQLTemp="select a.SN,a.Driver,a.DriverBirth,a.DriverID,a.DriverZip," &_
				"a.DriverAddress,a.DoubleCheckStatus,a.Rule1,a.DeallIneDate," &_
				"a.IllegalDate,a.IllegalAddress,a.BillNo,a.ForFeit1," &_
				"NVL(a.ForFeit2,0) ForFeit2,MemberStation," &_
				"(Select SendNumber from PasserSend where billsn=a.sn) SendNumber," &_
				"(Select AgentJob from PasserSend where billsn=a.sn) AgentJob," &_
				"(Select OpenGovNumber from PasserSend where billsn=a.sn) SendNo," &_
				"(Select SendDate from PasserSend where billsn=a.sn) SendDate," &_
				"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
				"(Select LimitDate from PasserSend where billsn=a.sn) LimitDate," &_
				"(Select SafeToExit from PasserSend where billsn=a.sn) SafeToExit," &_
				"(Select SafeAction from PasserSend where billsn=a.sn) SafeAction," &_
				"(Select SafeAssure from PasserSend where billsn=a.sn) SafeAssure," &_
				"(Select SafeDetain from PasserSend where billsn=a.sn) SafeDetain," &_
				"(Select SafeShutShop from PasserSend where billsn=a.sn) SafeShutShop," &_
				"(Select ForFeit from PasserSend where billsn=a.sn) SendFeit," &_
				"(Select AgentAddress from PasserJude where billsn=a.sn) AgentAddress," &_
				"(Select OpenGovNumber from PasserJude where billsn=a.sn) JudeNo," &_
				"(Select AgentName from PasserJude where billsn=a.sn) AgentName," &_
				"(Select AgentBirth from PasserJude where billsn=a.sn) AgentBirth," &_
				"(Select AgentID from PasserJude where billsn=a.sn) AgentID," &_
				"(Select JudeDate from PasserJude where billsn=a.sn) JudeDate," &_
				"(Select max(PayDate) from PasserPay where billsn=a.sn) PayDate," &_
				"' ' Administrative from PasserBase a where a.RecordStateID=0 and a.BillStatus<>9 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and Exists(select 'Y' from PasserSend where billsn=a.sn)"&orderstr

	If showCreditor Then
		
		strSQLTemp="select a.SN,a.Driver,a.DriverBirth,a.DriverID,a.DriverZip," &_
				"a.DriverAddress,a.DoubleCheckStatus,a.Rule1,a.DeallIneDate," &_
				"a.IllegalDate,a.IllegalAddress,a.BillNo,a.ForFeit1," &_
				"NVL(a.ForFeit2,0) ForFeit2,MemberStation," &_
				"(Select SendNumber from PasserSend where billsn=a.sn) SendNumber," &_
				"(Select AgentJob from PasserSend where billsn=a.sn) AgentJob," &_
				"(Select OpenGovNumber from PasserSend where billsn=a.sn) SendNo," &_
				"(Select SendDate from PasserSend where billsn=a.sn) SendDate," &_
				"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
				"(Select LimitDate from PasserSend where billsn=a.sn) LimitDate," &_
				"(Select SafeToExit from PasserSend where billsn=a.sn) SafeToExit," &_
				"(Select SafeAction from PasserSend where billsn=a.sn) SafeAction," &_
				"(Select SafeAssure from PasserSend where billsn=a.sn) SafeAssure," &_
				"(Select SafeDetain from PasserSend where billsn=a.sn) SafeDetain," &_
				"(Select SafeShutShop from PasserSend where billsn=a.sn) SafeShutShop," &_
				"(Select ForFeit from PasserSend where billsn=a.sn) SendFeit," &_
				"(Select AgentAddress from PasserJude where billsn=a.sn) AgentAddress," &_
				"(Select OpenGovNumber from PasserJude where billsn=a.sn) JudeNo," &_
				"(Select AgentName from PasserJude where billsn=a.sn) AgentName," &_
				"(Select AgentBirth from PasserJude where billsn=a.sn) AgentBirth," &_
				"(Select AgentID from PasserJude where billsn=a.sn) AgentID," &_
				"(Select JudeDate from PasserJude where billsn=a.sn) JudeDate," &_
				"(Select max(PayDate) from PasserPay where billsn=a.sn) PayDate," &_
				"Nvl((Select DeCode(AgentAddress,null,(select Administrative from zip where zipid=a.DriverZip),AgentAddress) from PasserSend where billsn=a.sn),'其它') Administrative" &_
				" from PasserBase a where a.RecordStateID=0 and a.BillStatus<>9 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and Exists(select 'Y' from PasserSend where billsn=a.sn) order by Administrative,Driver,Billno"

	End If 

	set rs=conn.execute(strSQLTemp)

	If sys_City="台中市" then

		Set UitObj = Server.CreateObject("Scripting.Dictionary")

		UitObj.Add "0410","204G02"

		UitObj.Add "0420","204G03"

		UitObj.Add "0430","204G04"

		UitObj.Add "0440","204G05"

		UitObj.Add "0450","204G06"

		UitObj.Add "0480","204G07"

		UitObj.Add "4A00","204H02"

		UitObj.Add "4C00","204H07"

		UitObj.Add "4E00","204H06"

		UitObj.Add "4D00","204H05"

		UitObj.Add "4F00","204H04"

		UitObj.Add "4H00","204H09"

		UitObj.Add "4B00","204H03"

		UitObj.Add "4G00","204H08"

		UitObj.Add "0460",""

		UitObj.Add "0406",""

	elseIf sys_City="台南市" then

		Set UitObj = Server.CreateObject("Scripting.Dictionary")

		UitObj.Add "7000","204N02"

		UitObj.Add "7100","204N03"

		UitObj.Add "7200","204N04"

		UitObj.Add "7300","204N05"

		UitObj.Add "7400","204N06"

		UitObj.Add "7500","204N07"

		UitObj.Add "A01","204O02"

		UitObj.Add "F01","204O03"

		UitObj.Add "D01","204O04"

		UitObj.Add "I01","204O05"

		UitObj.Add "G01","204O06"

		UitObj.Add "E01","204O07"

		UitObj.Add "H01","204O08"

		UitObj.Add "B01","204O09"

		UitObj.Add "J01","204O10"

		UitObj.Add "C01","204O11"

		UitObj.Add "0707",""

		UitObj.Add "0706",""
	end if
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style1 {font-size: 12px; }
.style2 {font-size: 12px;mso-style-parent:style0;mso-number-format:"\@";}
-->
</style>
</head>
<body>
<%
tmpAdministrative="A"
MemberStation=""
while Not rs.eof
	If tmpAdministrative <> trim(rs("Administrative")) Then
		
		If not ifnull(tmpAdministrative) and tmpAdministrative<>"A" Then Response.Write "<table border=0><tr><td></td></tr></table>"
		
		If sys_City="台中市" then

			If MemberStation = "" Then MemberStation=UitObj.Item(trim(rs("MemberStation")))
		end if
		
		tmpAdministrative=trim(rs("Administrative"))
		If not ifnull(tmpAdministrative) Then Response.Write "<table><tr><td>"&tmpAdministrative&"</td></tr></table>"
		%>
		<table border=1>
		 <tr>
		  <td class="style1">案件類別</td>
		  <td class="style1">移送案號</td>
		  <td class="style1">義務人</td>
		  <td class="style1">出生日期(義務人)</td>
		  <td class="style1">身分證字號(義務人)</td>
		  <td class="style1">職業(義務人)ex:士.農.工.商</td>
		  <td class="style1">郵遞區號(義務人)</td>
		  <td class="style1">戶籍地址(義務人)</td>
		  <td class="style1">法定代理人或代表人</td>
		  <td class="style1">出生日期(法定代理人或代表人)</td>
		  <td class="style1">身分證字號(法定代理人或代表人)</td>
		  <td class="style1">職業代碼(法定代理人或代表人)</td>
		  <td class="style1">郵遞區號(法定代理人或代表人)</td>
		  <td class="style1">戶籍地址(法定代理人或代表人)</td>
		  <td class="style1">營利事業統編</td>
		  <td class="style1">營業所</td>
		  <td class="style1">案由代碼</td>
		  <td class="style1">應納金額</td>
		  <td class="style1">行政處分或裁定確定日</td>
		  <td class="style1">義務發生之原因</td>
		  <td class="style1">義務發生之日期</td>
		  <td class="style1">義務發生之地點</td>
		  <td class="style1">發文字號</td>
		  <td class="style1">發文日期</td>
		  <td class="style1">保全措施(已限制出境)</td>
		  <td class="style1">保全措施(已禁止處分)</td>
		  <td class="style1">保全措施(已提供擔保)</td>
		  <td class="style1">保全措施(以假扣押)</td>
		  <td class="style1">保全措施(已勒令停業)</td>
		  <td class="style1">管理代號</td>
		  <td class="style1">本稅(健保費可內含)</td>
		  <td class="style1">罰鍰</td>
		  <td class="style1">短估金或老農津貼</td>
		  <td class="style1">教育經費或給付追回</td>
		  <td class="style1">滯報金/怠報金</td>
		  <td class="style1">核定補徵利息</td>
		  <td class="style1">行政救濟利息</td>
		  <td class="style1">滯納金</td>
		  <td class="style1">滯納期滿利息</td>
		  <td class="style1">其他費用或墊償費</td>
		  <td class="style1">執行憑證再移送</td>
		  <td class="style1">執行憑證編號</td>
		  <td class="style1">執行憑證核發日期</td>
		  <td class="style1">滯納金起算日</td>
		  <td class="style1">利息起算日</td>
		  <td class="style1">財產郵遞區號</td>
		  <td class="style1">財產地址</td>
		  <td class="style1">浮動滯納期滿利息</td>
		  <td class="style1">通訊郵遞區號（義務人）</td>
		  <td class="style1">通訊地址（義務人）</td>
		  <td class="style1">戶籍地電話（義務人）</td>
		  <td class="style1">通訊地電話（義務人）</td>
		  <td class="style1">通訊郵遞區號（法定代理人）</td>
		  <td class="style1">通訊地址（法定代理人）</td>
		  <td class="style1">戶籍地電話（法定代理人）</td>
		  <td class="style1">通訊地電話（法定代理人）</td>
		  <td class="style1">關稅滯納金計算迄日</td>
		  <td class="style1">關稅滯納金每日加徵金額勞保滯納金原應收金額</td>
		  <td class="style1">貨物稅、營業稅及菸酒稅滯納金計算起日</td>
		  <td class="style1">貨物稅、營業稅及菸酒稅滯納金計算迄日</td>
		  <td class="style1">貨物稅、營業稅及菸酒稅滯納金每二日加徵金額</td>
		  <td class="style1">貨物稅、營業稅及菸酒稅利息計算起日</td>
		  <td class="style1">貨物稅、營業稅及菸酒稅利息每日加徵金額</td>
		  <td class="style1">關稅利息每日加徵金額</td>
		  <td class="style1">獨資合夥</td>
		  <td class="style1">銷帳編號</td>
		  <td class="style1">繳款類別</td>
		  <td class="style1">繳納期間屆滿日</td>
		  <td class="style1">徵收期間屆滿日</td>
		  <td class="style1">性別</td>
		  <td class="style1">核銷機關(單位)名稱</td>
		  <td class="style1">核銷機關(單位) 統一編號</td>
		  <td class="style1">承辦機關(單位)名稱</td>
		  <td class="style1">立帳金融機構名稱</td>
		  <td class="style1">受款金融機構帳戶</td>
		  <td class="style1">帳號</td>
		  <td class="style1">限制出境日期</td>
		  <td class="style1">義務人羅馬拼音</td>
		  <td class="style1">法定代理人羅馬拼音</td>
		 </tr><%
	end if

		ZipID="":Sex=""
		ZipID=rs("DriverZip")
		MakeSureDate=rs("MakeSureDate")
		LimitDate=rs("LimitDate")
		
		If not IFnull(Trim(rs("DriverID"))) Then
			If Mid(Trim(rs("DriverID")),2,1)="1" Then
				Sex="男"
			elseif Mid(Trim(rs("DriverID")),2,1)="2" Then
				Sex="女"
			End if
		end If 

		paySum=0
		strSQL="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&rs("SN")
		set rspay=conn.execute(strSQL)
		If not rspay.eof Then paySum=cdbl(rspay("PaySum"))
		rspay.close


		Sys_ForFeit=cdbl(rs("ForFeit1"))+cdbl(rs("ForFeit2"))-cdbl(paySum)


		response.write "<tr>"
		response.write "<td class=""style1"">A</td>"
		response.write "<td class=""style1"">"&rs("SendNumber")&"</td>"
		response.write "<td class=""style1"">"&rs("Driver")&"</td>"
		response.write "<td class=""style2"">"&right("0"&gInitDT(rs("DriverBirth")),7)&"</td>"
		response.write "<td class=""style1"">"&rs("DriverID")&"</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1"">"&ZipID&"</td>"
		response.write "<td class=""style1"">"&rs("DriverAddress")&"</td>"
		response.write "<td class=""style1"">"&rs("AgentName")&"</td>"
		response.write "<td class=""style2"">"
			If trim(rs("AgentBirth")) <>""  Then
				Response.Write right("0"&gInitDT(rs("AgentBirth")),7)
			End if 
		Response.Write "</td>"
		response.write "<td class=""style1"">"&rs("AgentID")&"</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style2"">0601</td>"
		response.write "<td class=""style1"">"&Sys_ForFeit&"</td>"
		response.write "<td class=""style2"">"&right("0"&gInitDT(rs("MakeSureDate")),7)&"</td>"
		'response.write "<td class=""style1"">違反道路交通管理處罰條例第"&rs("Rule1")&"條</td>"
		response.write "<td class=""style1"">違反道路交通管理處罰條例</td>"
		response.write "<td class=""style2"">"&right("0"&gInitDT(rs("IllegalDate")),7)&"</td>"
		response.write "<td class=""style1"">"&rs("IllegalAddress")&"</td>"
		response.write "<td class=""style1"">"&BillPageUnit&"交字第"&rs("SendNo")&"號</td>"
		response.write "<td class=""style2"">"&right("0"&gInitDT(rs("SendDate")),7)&"</td>"
		response.write "<td class=""style1""></td>"
		if isnull(rs("SafeToExit")) then
			response.write "<td class=""style1""></td>"
		else
			response.write "<td class=""style1"">1</td>"
		end if
		if isnull(rs("SafeAction")) then
			response.write "<td class=""style1""></td>"
		else
			response.write "<td class=""style1"">1</td>"
		end if
		if isnull(rs("SafeDetain")) then
			response.write "<td class=""style1""></td>"
		else
			response.write "<td class=""style1"">1</td>"
		end if
		if isnull(rs("SafeShutShop")) then
			response.write "<td class=""style1""></td>"
		else
			response.write "<td class=""style1"">1</td>"
		end if
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1"">"&Sys_ForFeit&"</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1"">"
		If showCreditor Then
			strSQL="select OpenGovNumber from PasserCreditor where BillSN="&trim(rs("SN"))&" and PetitionDate is not null order by PetitionDate desc"

			set rsCre=conn.execute(strSQL)
			If not rsCre.eof Then
				Response.Write "1"
			End If 
			rsCre.close
		End if 
		Response.Write "</td>"
		response.write "<td class=""style1"">"
		If showCreditor Then
			strSQL="select OpenGovNumber from PasserCreditor where BillSN="&trim(rs("SN"))&" and PetitionDate is not null order by PetitionDate desc"

			set rsCre=conn.execute(strSQL)
			If not rsCre.eof Then
				Response.Write trim(rsCre("OpenGovNumber"))
			End If 
			rsCre.close
		End if 
		Response.Write "</td>"
		response.write "<td class=""style1"">"
		If showCreditor Then
			strSQL="select PetitionDate from PasserCreditor where BillSN="&trim(rs("SN"))&" and PetitionDate is not null order by PetitionDate desc"

			set rsCre=conn.execute(strSQL)
			If not rsCre.eof Then
				Response.Write right("0"&gInitDT(rsCre("PetitionDate")),7)
			End If 
			rsCre.close
		End if 
		Response.Write "</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1"">"&ZipID&"</td>"
		response.write "<td class=""style1"">"&rs("DriverAddress")&"</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style2"">"&right("0"&gInitDT(rs("LimitDate")),7)&"</td>"
		'宜蘭行政執行處說警局單位不需要填徵收期間屆滿日
		response.write "<td class=""style1"">"
'			if LimitDate(0)<>"00" then 
'			if  len(right("0"&LimitDate(0)&LimitDate(1)&LimitDate(2),7))>=7 then 	
'				if len(right("0"&LimitDate(0)&LimitDate(1)&LimitDate(2),7))<7 then
'					response.write ""
'				else
'					response.write right("0"&LimitDate(0)&LimitDate(1)&LimitDate(2),7)
'				end if
'			end if
'			end if
		Response.Write "</td>"
		response.write "<td class=""style1"">"&Sex&"</td>"
		response.write "<td class=""style1"">"&thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")&"</td>"
		response.write "<td class=""style2"">"&thePasserVATnumber&"</td>"
		response.write "<td class=""style1"">"&thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")&"</td>"
		response.write "<td class=""style1"">"&thePasserSendBankAccountName&"</td>"
		response.write "<td class=""style1"">"&thePasserSendBankName&"</td>"
		response.write "<td class=""style2"">"&thePasserSendBankAccount&"</td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "<td class=""style1""></td>"
		response.write "</tr>"

		rs.movenext

		If not rs.eof Then		
			If tmpAdministrative <> trim(rs("Administrative")) Then
				response.write "<tr>"
				response.write "<td class=""style1"">EOF</td>"
				response.write "<td class=""style1"">"&MemberStation&"</td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "<td class=""style1""></td>"
				response.write "</tr>"
			end if
		end if
	wend
	rs.close
	response.write "<tr>"
	response.write "<td class=""style1"">EOF</td>"
	response.write "<td class=""style1"">"&MemberStation&"</td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "<td class=""style1""></td>"
	response.write "</tr>"
	conn.close
 %>
</table>
</body>
</html>