<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
'基隆市
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
'	sys_City="基隆市"

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

If thenPasserUnitName="&nbsp;台中市交通警察大隊" Then thenPasserUnitName="臺中市政府警察局交通警察大隊"

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
		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,RecordDate,RecordMemberID,IllegalDate,Owner,OwnerAddress,OwnerZip,BillFillerMemberID from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"	
		if sys_City="台東縣" or sys_City="南投縣" then
			BilLBase=BilLBase&"  and billstatus<>'9'"
		End if
		'2012/05/4 南投陳淑雲說 監理單位已先入案 n 違規人已先繳結案 L ，不出來，固修改 c.Status in ('Y','S','n','L') 為  c.Status in ('Y','S') by jafe,目前只有南投有改，其他縣市未更新過去
		strSQL="select a.BillFillerMemberID,a.carno,a.BillNo,a.BillTypeID,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.MailchkNumber,a.Owner as BOwner,a.OwnerAddress as BOwnerAddress,a.OwnerZip as BOwnerZip from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and e.exchangetypeid=c.exchangetypeid and c.Status in ('Y','S','n','L') and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+)  and not exists(select 'Y' from dcicloseclosedata where billno=a.billno) "
		If sys_City="苗栗縣" Or sys_City="基隆市" Or sys_City="台中市" Then
			strSQL_Order=" order by c.OwnerZip,a.billno"
		Else
			strSQL_Order=" order by a.billno"
		End If 

' billstatus=2 and billtypeid=2
'		select count(*) from (
'select * from billbase a where billstatus>1 and recordstateid=0
'and illegaldate between to_date('2014/05/01','yyyy/mm/dd')
'and to_date('2014/06/01','yyyy/mm/dd')

' and  exists(select 'Y' from dcicloseclosedata where billno=a.billno)
')

		set rsfound=conn.execute(strSQL&strSQL_Order)
		While Not rsfound.eof
			filecmt=filecmt+1
			BillNo=rsfound("BillNo")&""
			BillFillerMemberID=rsfound("BillFillerMemberID")
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

		if sys_City = "南投縣" then
			strMem="select ChName,loginid from Memberdata where MemberID="&trim(rsfound("RecordMemberID"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				loginid = trim(rsMem("loginid"))
			end if
			rsMem.close
			set rsMem=nothing
		end if

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
					GetMailMan="&nbsp;"&funcCheckFont(trim(rsfound("Driver")),15,1)&"&nbsp;"
					GetMailAddress="&nbsp;"&trim(rsfound("DriverHomeZip"))&" "&ZipName2&trim(rsfound("DriverHomeAddress"))&"&nbsp;"
				else
					GetMailMan="&nbsp;"&funcCheckFont(trim(rsfound("Owner")),15,1)&"&nbsp;"
				GetMailAddress="&nbsp;"&trim(rsfound("OwnerZip"))&" "&ZipName&trim(rsfound("OwnerAddress"))&"&nbsp;"
				end if
			Else
			
				GetMailMan="&nbsp;"&funcCheckFont(trim(rsfound("Owner")),15,1)&"&nbsp;"
				GetMailAddress="&nbsp;"&trim(rsfound("OwnerZip"))&" "&ZipName&trim(rsfound("OwnerAddress"))&"&nbsp;"
				if sys_City="南投縣" And Trim(rsfound("BOwnerAddress"))<>"" Then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("BOwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing			
									
					GetMailMan=funcCheckFont(trim(rsfound("BOwner")),15,1)
					GetMailAddress=trim(rsfound("BOwnerZip"))&" "& ZipName2 &funcCheckFont(replace(replace(trim(rsfound("BOwnerAddress")),"臺","台"),ZipName2,""),15,1)
				End IF
			End if


'sys_City="南投縣"

			mailNumber=trim(replace(rsfound("MailchkNumber") &"" ,"000000",""))

			mailNumber=trim(replace(mailNumber &""," ",""))

		   if trim(mailNumber)="" Or trim(mailNumber)="0" then
        	   	mailNumber=trim(rsfound("mailNumber")) &""
           end if
	    
    	if sys_City="苗栗縣" Then
			mailNumber=Right("000000"&mailNumber,6)&"02982736000017"
		End If
		
    	if sys_City="南投縣" or sys_City="台中市" then
           mailNumber=trim(replace(mailNumber," ",""))

		   if trim(mailNumber)="" then
		   	mailNumber=trim(rsfound("mailNumber")) &""
				  if mailNumber<>"" then
    			   for j=1 to 14-len(trim(mailNumber))
			     		mailnumber="0" & mailnumber 
				   next 		
				  end If
		   Else
			if sys_City="南投縣" then 
		   	mailNumber=trim(rsfound("mailNumber")) &""
				   for j=1 to 14-len(trim(mailNumber))
			     		mailnumber="0" & mailnumber 
				   next 				
			End if
		   end If
		   
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
			Sys_BillNo_BarCode=BillNo
           '	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
			If sys_City="高雄市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,57,160,1
			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,57,160
			end If
			
%>

<div id="R1" style="position:relative;">
<table border="0" width="100" id="table1" height="625" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<table border="0" width="100" id="table2" cellspacing="0" cellpadding="0" height="625">
			<tr>
			<td colspan="3" align="left">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font face="標楷體" size="2">
			<%

			If sys_City="基隆市" then
				response.write "序號:" & filecmt & "&nbsp;&nbsp;"
			End if
			%>
			<%If sys_City="台中市" Then%>
									<font face="標楷體" size="5">
			一般
			<%elseIf sys_City<>"南投縣" Then%>
						<font face="標楷體" size="5">
			傳真

			<%End if%>
			查詢國內各類掛號郵件查單</font>
			<%

			If sys_City="基隆市" then
				response.write "&nbsp;列印日期:"&Year(date())-1911&"/"&Month(Date())&"/"&Day(Date())&"&nbsp;"
			Else
				response.write "　　　　　　　　"
			End if
			%>
			<font face="標楷體">編列第　　　　　　　號
			<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg">
			<%If sys_City="台中市" Then
					If BillFillerMemberID<>"" Then
					   					   					   strUnit="select b.UnitTypeID,b.UnitName from memberdata a,UnitInfo b where  a.UnitID=b.UnitID and a.memberid=" & BillFillerMemberID
					   Set rstmp=conn.execute(strUnit)
					   If Not rstmp.eof Then 
					            strUnit="select  UnitName from UnitInfo where UnitID='" & rstmp("UnitTypeID") &"'"
					            Set rstmp2=conn.execute(strUnit)
										response.write Replace(replace(rstmp2("UnitName") & rstmp("UnitName"),rstmp2("UnitName") & rstmp2("UnitName"),rstmp2("UnitName")),"交通警察大隊","")
										Set rstmp2=nothing
					  end If
					  Set rstmp=nothing
				end if
			End if%>
			</font><tr>
				<td width="485" align="left" valign="top">
				<table border="1" width="485" id="table3" cellspacing="0" cellpadding="0" height="625">
					<tr>
						<td rowspan="3" width="16" align="center">
						<font face="標楷體">受理局填寫</font></td>
						<td rowspan="2" width="60" colspan="2" align="center">
						<font face="標楷體">原　寄<br>局　名</font></td>
						<%If sys_City="基隆市" Then%>
							<%If Session("Unit_ID")="0220" then%>
								<td width="74"  rowspan="2">基隆仁二路</td>
							<%else%>
								<td width="74"  rowspan="2">　　　</td>
							<%End if%>
						<%ElseIf sys_City="苗栗縣" Then%>
								<td width="74"  rowspan="2">中苗郵局</td>
						<%ElseIf sys_City="台中市" Then%>
								<td width="74"  rowspan="2">臺中<br><font size="2">民權路郵局</font></td>
						<%Else%>
							<td width="74" rowspan="2">　　　</td>
						<%End if%>
						<td colspan="20" align="center"><font face="標楷體">條碼掛號(填寫完整 14 或 20 碼)</font></td>
					</tr>
					<tr>
						<td colspan="6" align="center"><font face="標楷體">掛號號碼</font></td>
						<td  align="center" colspan="6"><font face="標楷體">原寄局碼</font></td>
						<td  align="center" colspan="2"><font size="2" face="標楷體">郵件別</font></td>
						<td  align="center" colspan="5"><font face="標楷體">寄達局碼</font></td>
						<td  align="center"><font face="標楷體">檢</font></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="44" align="center">
						<font face="標楷體">掛　號<br>種　類</font></td>
						<%If sys_City="基隆市" Then%>
							<%If Session("Unit_ID")="0220" then%>
								<%If mid(mailNumber,13,2)="36" Then %>
									<td width="74" height="44">雙掛</td>
								<%else%>
									<td width="74" height="44">　　　</td>
								<%End if%>
							<%else%>
								<td width="74" height="44">　　　</td>
							<%End if%>
						<%else%>
							<td width="74" height="44">　　　</td>
						<%End if%>

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
						<td width="13" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,12,1)<>"" then response.write mid(mailNumber,12,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,13,1)<>"" then response.write mid(mailNumber,13,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,14,1)<>"" then response.write mid(mailNumber,14,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,15,1)<>"" then response.write mid(mailNumber,15,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,16,1)<>"" then response.write mid(mailNumber,16,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,17,1)<>"" then response.write mid(mailNumber,17,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,18,1)<>"" then response.write mid(mailNumber,18,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,19,1)<>"" then response.write mid(mailNumber,19,1) else response.write "&nbsp;"%></font></td>
						<td width="16" height="44" align="center"><font face="標楷體"><%if mid(mailNumber,20,1)<>"" then response.write mid(mailNumber,20,1) else response.write "&nbsp;"%></font></td>
					</tr>
					<tr>
						<td width="16" rowspan="6" align="center">
						<font face="標楷體">查</font><p><br><font face="標楷體">詢</font></p>
						<p>&nbsp;<br><font face="標楷體">人</font></p>
						<p>&nbsp;<br><font face="標楷體">填</font></p>
						<p><br><font face="標楷體">寫</font></td>
						<td width="60" colspan="2" height="36" align="center">
						<font face="標楷體">交　寄<br>日　期</font></td>
						<td colspan="21" height="36"><font face="標楷體">　　　<%
						If s_Year<>"0" Then response.write s_Year
						%>　年　<%
						If s_month<>"0" Then response.write s_month
						%>　月　<%
						If s_day<>"0" Then response.write s_day
						%>　日　<%
						If s_hour<>"00" and s_hour<>"0" Then response.write s_hour
						%>　時</font></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="50" align="center">
						<font face="標楷體">報　值<br>保　價
						<br>金　額</font></td>
						<td width="124" height="50" colspan="4">　　</td>
						<td height="50" width="47" align="center" colspan="3"><font face="標楷體">重量</font></td>
						<td height="50" width="90" align="center" colspan="6">　</td>
						<td height="50" width="36" align="center" colspan="2"><font face="標楷體">內裝</font></td>
						<td height="50" width="96" colspan="6" align="center">&nbsp;
						<%If sys_City="南投縣" Then
							response.write loginid&"-1"
						  End if
						%>
						</td>
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
						<font face="標楷體" size="2">違規單號</font></td>
					</tr>
					<tr>
						<td width="96" colspan="6" height="40">&nbsp;<%=BillNo%></td>
					</tr>
					<tr>
						<td width="21" align="center" height="63"><font face="標楷體">收件人</font></td>
						<td width="37" align="center" height="63"><font face="標楷體">姓名地址電話</font></td>
						<td width="401" colspan="21" height="63"><font face="標楷體">
							<%=(GetMailMan)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%>
							<%If sys_City="南投縣" Then  response.write CarNo%>
							<Br>
							<%=(GetMailAddress)%></font>
						</td>
					</tr>
					<tr>
						<td width="60" colspan="2" align="center">
						<font face="標楷體">查　詢<br>結　果</font></td>
						<%If sys_City="基隆市" Then%>
							<%If Session("Unit_ID")="0220" then%>
							<td width="401" colspan="21"><font face="標楷體">　□電話通知　　<span lang=EN-US style='font-family:"Wingdings 2";
mso-ascii-font-family:"Times New Roman";mso-hansi-font-family:"Times New Roman";
mso-char-type:symbol;mso-symbol-font-family:"Wingdings 2"'><span
style='mso-char-type:symbol;mso-symbol-font-family:"Wingdings 2"'>R</span></span>傳真　　<span lang=EN-US style='font-family:"Wingdings 2";
mso-ascii-font-family:"Times New Roman";mso-hansi-font-family:"Times New Roman";
mso-char-type:symbol;mso-symbol-font-family:"Wingdings 2"'><span
style='mso-char-type:symbol;mso-symbol-font-family:"Wingdings 2"'>R</span></span>補發回執</font></td>
							<%else%>
								<td width="401" colspan="21"><font face="標楷體">　□電話通知　　□傳真　　□補發回執</font></td>
							<%End if%>
						<%else%>
								<td width="401" colspan="21"><font face="標楷體">　□電話通知　　□傳真　　□補發回執</font></td>
						<%End if%>
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
									<tr><br>
										<td width="25" ><font face="標楷體">　　</font><br>　</td>
										<td width="433" align="left" valign="top">
										<table border="1"  id="table6" cellspacing="0" cellpadding="0" height="72" width="183">
											<tr>
												<td width="179" height="30">
												<font face="標楷體">除快捷郵件外，其他郵件應收傳真費，用郵票或郵資券粘貼於此。</font></td>
											</tr>
										</table>
										</td>
									</tr>
									<tr>
										<td colspan="2" height="132">
										<font face="標楷體"><br>　查右列郵件，據寄件人聲稱，並未寄到，請即迅為查詢見覆。<br>
										　本局傳真號碼「　　　　　　　　　　　　　　」。
										<br>　
										<br>　　　　　　　　　　　　　經辦員：
										<br>　　　　　　　　　　　　　主　管：
										<br>　中華民國　　　　年　　　　月　　　　日</font></td>
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
<div  style="position:absolute;left:920px;top:410px"><img src="../Image/MailPic2.JPG" width="90" height="90" /></div>
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
printWindow(false,7,10.08,5.08,0);
</script>
</html>