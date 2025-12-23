<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

chkDate=trim(request("chkDate"))
strDate=split("BillFillDate,IllegalDate,RecordDate",",")
strDateName=split("填單日,違規日,建檔日",",")
UserId = Session("User_ID")
startDate_q = Trim(Request("startDate_q"))
endDate_q = Trim(Request("endDate_q"))
unit = Request("unit")
UnitID_q = Request("UnitID_q")
unitList=trim(request("unitSelectlist"))
Batchnumber_q=trim(request("Batchnumber"))
Memlist_q=trim(request("MemSelectlist"))
Server.ScriptTimeout=86400


thenPasserUnit=""
strSQL="select UnitID,UnitTypeID,UnitLevelID,UnitName,Address,TEL from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitID2=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
    thenPasserUnitName="&nbsp;"&sys_City&trim(rsunit("UnitName"))
    thenPasserUnitAddress="&nbsp;"&trim(rsunit("Address"))
	thenPasserUnitTel="&nbsp;"&trim(rsunit("TEL"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
elseif Sys_UnitLevelID=2 and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
end if
set rsunit=conn.Execute(strSQL)
if Not rsunit.eof then Sys_UnitID=trim(rsunit("UnitID"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

tmpSql=""
'入案批號
if Batchnumber_q<>"" then
	tmpSql = tmpSql & " and SN in (select BillSn from Dcilog where BatchNumber='" & Batchnumber_q & "')"
end if
'建檔人員
if Memlist_q<>"" then
	tmpSql = tmpSql & " and RecordMemberId in (" & Memlist_q & ")"
end if
'統計日期
if startDate_q<>"" then
	tmpSql = tmpSql & " and "&strDate(chkDate)&" Between To_Date('" & gOutDT(startDate_q)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(endDate_q)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
end if
'舉發單號
if trim(request("startBillNo_q"))<>"" then
	tmpSql = tmpSql & " and BillNo Between '" & trim(request("startBillNo_q")) & "' And '" & trim(request("endBillNo_q")) & "'"
end if
'舉發單位
If unit="y" Then
	unitList = Split(unitList,",")
	Sys_UnitID=""
	for i=0 to UBound(unitList)
		if Sys_UnitID<>"" then Sys_UnitID=Sys_UnitID&"','"
		Sys_UnitID=Sys_UnitID&unitList(i)
	next
	UnitSql = " and BillUnitID in ('" & Sys_UnitID & "')"
End If

P_UnitName=thenPasserCity
strSQL="select UnitName from UnitInfo where UnitID='"&UnitID_q&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then P_UnitName=trim(rsunit("UnitName"))
rsunit.close
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>受理局填寫</title>

<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<%
filecmt=0
		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,RecordDate,RecordMemberID,IllegalDate from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"		
		if sys_City="台東縣" then
			BilLBase=BilLBase&"  and billstatus<>'9'"
		End if	
		strSQL="select a.BillNo,a.BillTypeID,a.CarNo,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.MailchkNumber from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and c.Status in ('Y','S','n','L') and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) order by a.billno"
		set rsfound=conn.execute(strSQL)
		While Not rsfound.eof
ZipName2=""
			filecmt=filecmt+1
			BillNo=rsfound("BillNo")&""
			CarNo=rsfound("CarNo")&""
			s_date=gInitDT(trim(rsfound("RecordDate")))
			s_hour=right("0"&hour(rsfound("RecordDate")),2)
			s_minute=right("0"&minute(rsfound("RecordDate")),2)
			RecordDate=s_date&"<br>"&s_hour&s_minute
			s_Year=year(trim(rsfound("RecordDate")))-1911
			s_Month=right("0"&month(trim(rsfound("RecordDate"))),2)
			s_Day=right("0"&day(trim(rsfound("RecordDate"))),2)

			s_date=gInitDT(trim(rsfound("IllegalDate")))
			s_hour=right("0"&hour(rsfound("IllegalDate")),2)
			s_minute=right("0"&minute(rsfound("IllegalDate")),2)
			IllegalDate=s_date&"<br>"&s_hour&s_minute
			s_Year=year(trim(rsfound("IllegalDate")))-1911
			s_Month=right("0"&month(trim(rsfound("IllegalDate"))),2)
			s_Day=right("0"&day(trim(rsfound("IllegalDate"))),2)

			s_date=gInitDT(trim(rsfound("mailDate")))
			s_hour=right("0"&hour(rsfound("mailDate")),2)
			s_minute=right("0"&minute(rsfound("mailDate")),2)
			mailDate=s_date
			s_Year=year(trim(rsfound("mailDate")))-1911
			s_Month=right("0"&month(trim(rsfound("mailDate"))),2)
			s_Day=right("0"&day(trim(rsfound("mailDate"))),2)
			'&"<br>"&s_hour&s_minute

	    	if sys_City="金門縣" or sys_City="澎湖縣"  then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if

	    	if sys_City="金門縣" or sys_City="澎湖縣"  then
				ZipName2=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if


			If trim(rsfound("BillTypeID"))="1" Then
				if trim(rsfound("DriverHomeZip"))<>"" and not isnull(rsfound("DriverHomeZip")) then
					GetMailMan="&nbsp;"&trim(rsfound("Driver"))&"&nbsp;"
					GetMailAddress="&nbsp;"&trim(rsfound("DriverHomeZip"))&" "&ZipName2&trim(rsfound("DriverHomeAddress"))&"&nbsp;"
				else
					GetMailMan="&nbsp;"&trim(rsfound("Owner"))&"&nbsp;"
				GetMailAddress="&nbsp;"&trim(rsfound("OwnerZip"))&" "&ZipName&trim(rsfound("OwnerAddress"))&"&nbsp;"
				end if
			else
					GetMailMan="&nbsp;"&trim(rsfound("Owner"))&"&nbsp;"
				GetMailAddress="&nbsp;"&trim(rsfound("OwnerZip"))&" "&ZipName&trim(rsfound("OwnerAddress"))&"&nbsp;"
			End if


'sys_City="南投縣"

           mailNumber=trim(replace(trim(rsfound("MailchkNumber")) &""," ",""))

		   if trim(mailNumber)="" Or trim(mailNumber)="0" then
        	   	mailNumber=trim(rsfound("mailNumber")) &""
           end if

    	if sys_City="南投縣" or sys_City="台中市" then
           mailNumber=trim(replace(trim(rsfound("MailchkNumber")) &""," ",""))

		   if trim(mailNumber)="" then
		   	mailNumber=trim(rsfound("mailNumber")) &""
				  if mailNumber<>"" then
    			   for j=1 to 14-len(trim(mailNumber))
			     		mailnumber="0" & mailnumber 
				   next 		
				  end if
		   end if
       end if
		if sys_City="花蓮縣" and Sys_UnitID2="B000" then
			mailNumber=""
			s_Year=""
			s_Month=""
			s_Day=""
			s_hour=""
			BillNo=""
		end if
			if cint(filecmt)>1 then response.write "<div class=""PageNext"">&nbsp;</div>"
           	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
%>

<div id="R1" style="position:relative;">
<table border="0" width="100" id="table1" height="625" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<table border="0" width="100" id="table2" cellspacing="0" cellpadding="0" height="625">
			<tr>
			<td colspan="3" align="right">
			<font face="標楷體" size="5">傳真查詢國內各類掛號郵件查單</font>　　　　　　　　<font face="標楷體">編列第　　　　　　　號&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			</font>&nbsp;<tr>
				<td width="485" align="left" valign="top">
				<table border="1" width="485" id="table3" cellspacing="0" cellpadding="0" height="625">
					<tr>
						<td rowspan="3" width="16" align="center">
						<font face="標楷體">受理局填寫</font></td>
						<td rowspan="2" width="80" colspan="2" align="center">
						<font face="標楷體">原　寄<br>局　名</font></td>
						<td width="74" rowspan="2"  align="center"><font face="標楷體">
						<%if session("Unit_ID")="A000" then %>
						花蓮市府前路郵局
						<%else%>
						&nbsp;
						<%End if%></font></td>
						<td colspan="20" align="center"><font face="標楷體">條&nbsp;碼&nbsp;掛&nbsp;號&nbsp;收&nbsp;據&nbsp;之</font></td>
					</tr>
					<tr>
						<td colspan="6" align="center"><font face="標楷體">掛號號碼</font></td>
						<td  align="center" rowspan="2"><font size="2" face="標楷體">&nbsp;</font></td>
						<td  align="center" colspan="12"><font face="標楷體">&nbsp;&nbsp;原&nbsp;&nbsp;寄&nbsp;&nbsp;局&nbsp;&nbsp;碼&nbsp;&nbsp;</font></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="44" align="center">
						<font face="標楷體">掛　號<br>種　類</font></td>
						<%if len(mailNumber)<=6 then %>
						<td width="74" height="44"  align="center"><font face="標楷體">雙掛號</font></td>
						<td height="44" width="14" align="center"><font face="標楷體"><%if mid(mailNumber,1,1)<>"" then response.write mid(mailNumber,1,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,2,1)<>"" then response.write mid(mailNumber,2,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,3,1)<>"" then response.write mid(mailNumber,3,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="13" align="center"><font face="標楷體"><%if mid(mailNumber,4,1)<>"" then response.write mid(mailNumber,4,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,5,1)<>"" then response.write mid(mailNumber,5,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,6,1)<>"" then response.write mid(mailNumber,6,1) else response.write "&nbsp;"%></font></td>

						<td width="15" height="44" align="center"><font face="標楷體">9</font></td>
						<td width="14" height="44" align="center"><font face="標楷體">7</font></td>
						<td width="13" height="44" align="center"><font face="標楷體">3</font></td>
						<td width="13" height="44" align="center"><font face="標楷體">0</font></td>
						<td width="12" height="44" align="center"><font face="標楷體">0</font></td>
						<td width="13" height="44" align="center" colspan="2"><font face="標楷體">7</font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體">1</font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體">7</font></td>
						<%else%>
						<td width="74" height="44"  align="center"><font face="標楷體">雙掛號</font></td>
						<td height="44" width="14" align="center"><font face="標楷體"><%if mid(mailNumber,1,1)<>"" then response.write mid(mailNumber,1,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,2,1)<>"" then response.write mid(mailNumber,2,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,3,1)<>"" then response.write mid(mailNumber,3,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="13" align="center"><font face="標楷體"><%if mid(mailNumber,4,1)<>"" then response.write mid(mailNumber,4,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,5,1)<>"" then response.write mid(mailNumber,5,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(mailNumber,6,1)<>"" then response.write mid(mailNumber,6,1) else response.write "&nbsp;"%></font></td>

						<td width="15" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,7,1)<>"" then response.write mid(mailNumber,7,1) else response.write "&nbsp;"%></font></td>
						<td width="14" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,8,1)<>"" then response.write mid(mailNumber,8,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,9,1)<>"" then response.write mid(mailNumber,9,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,10,1)<>"" then response.write mid(mailNumber,10,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,11,1)<>"" then response.write mid(mailNumber,11,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(mailNumber,12,1)<>"" then response.write mid(mailNumber,12,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(mailNumber,13,1)<>"" then response.write mid(mailNumber,13,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(mailNumber,14,1)<>"" then response.write mid(mailNumber,14,1) else response.write "&nbsp;"%></font></td>
						<%end if%>
					</tr>
					<tr>
						<td width="16" rowspan="6" align="center">
						<font face="標楷體">查</font><p><br><font face="標楷體">詢</font></p>
						<p>&nbsp;<br><font face="標楷體">人</font></p>
						<p>&nbsp;<br><font face="標楷體">填</font></p>
						<p><br><font face="標楷體">寫</font></td>
						<td width="60" colspan="2" height="36" align="center">
						<font face="標楷體">交　寄<br>日　期</font></td>
						<td colspan="21" height="36"><font face="標楷體">　<%=s_Year%>　年　<%=s_month%>　月　<%=s_day%>　日　</font><img src="../BarCodeImage/<%=BillNo%>.jpg"></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="50" align="center">
						<font face="標楷體">報　值<br>保　價
						<br>金　額</font></td>
						<td width="110" height="50" colspan="4">　　</td>
						<td height="50" width="47" align="center" colspan="3"><font face="標楷體">重量</font></td>
						<td height="50" width="90" align="center" colspan="6">　</td>
						<td height="50" width="36" align="center" colspan="2"><font face="標楷體">內裝</font></td>
						<td height="50" width="96" colspan="6" align="center">　</td>
					</tr>
					<tr>
						<td width="21" rowspan="2" align="center">
						<font face="標楷體">寄件人</font></td>
						<td width="37" rowspan="2" align="center">
						<font face="標楷體">姓名住址電話</font></td>
						<td width="299" colspan="15" rowspan="2"><font face="標楷體">
							<%=thenPasserUnitName%>
							<br>
							<%=thenPasserUnitAddress%>
							<br>
							<%=thenPasserUnitTel%></font>
						</td>
						<td width="96" colspan="6" height="26" align="center">
						<font face="標楷體" size="2">寄件人FAX號碼</font></td>
					</tr>
					<tr>
						<td width="96" colspan="6" height="40">&nbsp;<%=BillNo%></td>
					</tr>
					<tr>
						<td width="21" align="center" height="63"><font face="標楷體">收件人</font></td>
						<td width="37" align="center" height="63"><font face="標楷體">姓名地址電話</font></td>
						<td width="401" colspan="21" height="63"><font face="標楷體">
						&nbsp;清冊編號&nbsp;<%=filecmt%>
						<br>
							<%=funcCheckFont(GetMailMan,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%>
							<%If sys_City="南投縣" Then  response.write CarNo%>
							<Br>
							<%=funcCheckFont(replace(replace(GetMailAddress,"臺","台"),ZipName&ZipName,ZipName),16,1)%></font>
						</td>
					</tr>
					<tr>
						<td width="60" colspan="2" align="center">
						<font face="標楷體">查　詢<br>結　果</font></td>
						<td width="401" colspan="21"><font face="標楷體">　□電話通知　　▉傳真　　□補發回執</font></td>
					</tr>
					<tr>
						<td width="16" align="center" rowspan="2"><br><font face="標楷體">受</font><p>&nbsp;<br><font face="標楷體">理</font></p>
						<p>&nbsp;<br><font face="標楷體">局</font></p>
						<p>&nbsp;<br><font face="標楷體">填</font></p>
						<p>&nbsp;<br><font face="標楷體">寫</font></td>
						<td width="60" colspan="2" align="center">
						<font face="標楷體">投　遞<br>局　別</font></td>
						<td width="401" colspan="21"><font face="標楷體">　　　　　　　　　郵局</font></td>
					</tr>
					<tr>
						<td width="95" colspan="23"  height="82">

								<table border="0" width="457" id="table5" height="204" cellspacing="0" cellpadding="0">

									<tr>
										<td colspan="2" height="132">
										<font face="標楷體"><br>　查右列郵件，據寄件人聲稱，並未寄到，請即迅為查詢見覆。<br>
										　本局傳真號碼「&nbsp;03-8344682&nbsp;」。
										<br>　
										<br>　　　　　　　　　　　　　經辦員：
										<br>　　　　　　　　　　　　　主　管：
										<br>　中華民國　　　　年　　　　月　　　　日</font></td>
										<tr>
										<td width="433" align="right" valign="top">
										<table border="1"  id="table6" cellspacing="0" cellpadding="0" height="72" width="183">
											<tr>
												<td width="179" height="30">
												<font face="標楷體">除快捷郵件外，其他郵件應收傳真費，用郵票或郵資券粘貼於此。</font></td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
						</table>
						</td>

				</td>
				
				<td align="left" valign="top"><font color=#ffffff>=</font></td>
				
				<td width="528" align="left" valign="top">
				<table border="1" width="528" id="table7" height="639" cellspacing="0" cellpadding="0">
					<tr>
						<td height="84" width="29" align="center">
						<font face="標楷體">投<br>遞
						<br>局
						<br>(一)</font></td>
						<td height="84" width="483">
						<font face="標楷體">該件於　　年　　月　　日隨第　　號清單第　　頁第　　格　　發<br>往　　貴局投遞(招領)請
						詳查
						<br>　　　年　　月　　日　　　　郵局　經辦員：
						<br>　　　　　　　　　　　　　　　　　主　管：</font></td>
					</tr>
					<tr>
						<td height="78" width="29" align="center">
						<font face="標楷體">投<br>遞
						<br>局
						<br>(二)</font></td>
						<td height="78" width="483"><font face="標楷體">該件於　　年　　月　　日隨第　　號清單第　　頁第　　格　　發<br>往　　貴局投遞(招領)請
						詳查
						<br>　　　年　　月　　日　　　　郵局　經辦員：
						<br>　　　　　　　　　　　　　　　　　主　管：</font></td>
					</tr>
					<tr>
						<td width="29" height="272" align="center">
						<font face="標楷體">投</font><p><font face="標楷體"><br>遞</font></p>
						<p><font face="標楷體">&nbsp;<br>局</font></p>
						<p><font face="標楷體">&nbsp;<br>(三)</font></td>
						<td height="272" width="483"><font face="標楷體">
						茲將最後查得結果說明如下（V）：</font><p><font face="標楷體">
						□一、查該件業於　　年　　月　　日妥投，妥投收據傳真如後，以為投到之據。</font></p>
						<p><font face="標楷體">□二、該件未投遞，原因如左：</font></p>
						<p><font face="標楷體">查該件</font></p>
						<p><font face="標楷體">　　　　　　　　　　　　　　　　經辦員：</font></p>
						<p><font face="標楷體">　　　　　　　　　　　　　　　　主　管：</font></p>
						<p><font face="標楷體">中華民國　　　　　年　　　　　月　　　　　日</font></td>
					</tr>
					<tr>
						<td colspan="2" align="center">
						<table border="0" width="400" id="table8" cellspacing="0" cellpadding="0">
							<tr>
								<td width="97">　</td>
								<td><font face="標楷體">妥投收據(或影本)貼此處</font><p>
								<font face="標楷體">
						一併傳真至原查詢局後，
						</font>
								<p><font face="標楷體">
						收據仍取下存檔。</font></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td colspan="2" height="35"><font face="標楷體">　補到回執已收訖：寄件人簽章</font></td>
					</tr>
				</table>
				　</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
</div>
<%			
response.flush
rsfound.movenext
		Wend%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,7,10.08,5.08,0);
</script>
</html>