<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印--A4 size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
'on Error Resume Next

Server.ScriptTimeout=6000

PBillSN=split(trim(request("BillSN")),",")

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

for i=0 to Ubound(PBillSN)

if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
	
	strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

strSQL="select count(1) cnt from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))
set rsSend=conn.execute(strSQL)

If cdbl(rsSend("cnt")) = 0 Then

	strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,MailDate,SendMailStation,ArriveType,ReturnResonID,Note) values((select nvl(Max(SN),0)+1 as cnt from PassersEndArrived),"&trim(rsbil("BillSN"))&",sysdate,"&Session("User_ID")&","&Session("User_ID")&",sysdate,null,2,null,null)"

	conn.execute(strSQL)

End if 
rsSend.close

strSql="select * from PasserBase where SN="&trim(rsbil("BillSN"))

set rs=conn.execute(strSql)

Sys_BillUnitID="":Sys_RecordMemberID="":Sys_BillFillerMemberID=""
Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip=""
Sys_Illegaladdress="":Sys_IllegalSpeed="":Sys_RuleSpeed="":Sys_Note=""
Sys_Rule1="":Sys_Rule2="":Sys_Level1=0:Sys_Level2=0:Sum_Level=0
Sys_RuleVer="":Sys_MailDate=""

if Not rs.eof then 

	Sys_BillUnitID=trim(rs("BillUnitID"))
	Sys_RecordMemberID=trim(rs("RecordMemberID"))
	Sys_BillFillerMemberID=trim(rs("BillFillerMemberID"))
	Sys_Owner=trim(rs("Driver"))
	Sys_OwnerAddress=trim(rs("DriverAddress"))
	Sys_OwnerZip=trim(rs("DriverZip"))
	Sys_Illegaladdress=trim(rs("ILLEGALADDRESS"))
	Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
	Sys_RuleSpeed=trim(rs("RuleSpeed"))
	Sys_Note=trim(rs("Note"))
	Sys_RuleVer=trim(rs("RuleVer"))

	Sys_DCIReturnStation=trim(rs("MemberStation"))

	Sys_BillNo=trim(rs("BillNo"))
	Sys_CarNo=trim(rs("CarNo"))

	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))

	If isnumeric(rs("FORFEIT1")) Then Sys_Level1=rs("FORFEIT1")
	If isnumeric(rs("FORFEIT2")) Then Sys_Level2=rs("FORFEIT2")

	Sum_Level=Sys_Level1+Sys_Level2

	Sys_DCIRETURNCARTYPE="微型電動二輪車"

	Sys_MailDate=trim(rs("BillFillDate"))


end If 

if Not rs.eof then
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end If 

if Not rs.eof then
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_h=""
end If 

if Not rs.eof then
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_m=""
end If 

if Not rs.eof then
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if

if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end If 



rs.close

If ifnull(Sys_OwnerAddress) Then
	strSQL="select Owner,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and Status='S'"
	
	set rsfi=conn.execute(strSql)

	if Not rsfi.eof then
		If Not ifnull(trim(rsfi("Owneraddress"))) Then

			Sys_Owner=trim(rs("Owner"))
			Sys_OwnerAddress=trim(rs("OwnerZip"))
			Sys_OwnerZip=trim(rs("Owneraddress"))
			
			strSQL="update passerbase set" & _
			" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
			",DriverAddress='"&trim(rsfi("Owneraddress"))&"(車)'" & _
			" where SN="&trim(rsbil("BillSN"))

			conn.execute(strSQL)
		end if
	end If 
	rsfi.close
end If 

	strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
	rszip.close

	Sys_OwnerAddress=replace(replace(Sys_OwnerAddress,"臺","台"),Sys_OwnerZipName,"")

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,c.Content from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.jobid=c.id(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
if Not mem.eof then Sys_BillFillerJobName=mem("Content")
mem.close

strSQL="select UnitName,Tel,Address from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysAddress=Unit("Address")
Unit.close


Sys_IllegalRule1=""

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then

	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end If 

Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end If 

Sys_STATIONNAME="":Sys_StationTel="":StationID="":theBankAccount="":theBankName=""

strSql="select UnitID,UnitName,Tel,BankName,BankAccount from Unitinfo where unitid='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_STATIONNAME=trim(rs("UnitName"))
if Not rs.eof then Sys_StationTel=trim(rs("Tel"))
if Not rs.eof then Sys_StationID=trim(rs("UnitID"))
if Not rs.eof then theBankAccount=trim(rs("BankAccount"))
if Not rs.eof then theBankName=trim(rs("BankName"))
rs.close

Sys_MailNumber=0

DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

rsbil.close

%>
<div id="L78" style="position:relative;"><%
%>
<div id="Layer2" style="position:absolute; left:30px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>

<div id="Layer4" style="position:absolute; left:150px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<!--<div id="Layer5" style="position:absolute; left:185px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:625px; top:25px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:625px; top:35px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:625px; top:45px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:5px; top:70px; width:202px; height:36px; z-index:5"><%
	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
%></div>
<!--<div id="Layer42" style="position:absolute; left:250px; top:70px; width:202px; height:36px; z-index:5"><%="<font size=1>"&SysUnit&"<br>("&SysUnitTel&")</font>"%></div>-->
<div id="Layer10" style="position:absolute; left:500px; top:65px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:485px; top:110px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer12" style="position:absolute; left:100px; top:125px; width:150px; height:11px; z-index:3">逕行舉發</div>
<div id="Layer15" style="position:absolute; left:250px; top:145px; width:100px; height:10px; z-index:8"><font size=2><%'if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"　"&right("0"&Sys_DriverBirth(1),2)&"　"&right("0"&Sys_DriverBirth(2),2)%></font></div>
<div id="Layer16" style="position:absolute; left:420px; top:145px; width:106px; height:13px; z-index:9"><%'=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:610px; top:145px; width:99px; height:12px; z-index:10"><%'=fastring%></div>
<div id="Layer18" style="position:absolute; left:100px; top:170px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:250px; top:170px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:490px; top:170px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,22,1)%></div>
<div id="Layer21" style="position:absolute; left:100px; top:200px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,22,1)%></div>

<div id="Layer22" style="position:absolute; left:100px; top:220px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" style="position:absolute; left:155px; top:220px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" style="position:absolute; left:200px; top:220px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:250px; top:220px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" style="position:absolute; left:300px; top:220px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" style="position:absolute; left:390px; top:225px; width:600px; height:31px; z-index:20"><%
	response.write "<font size=3>"
	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if
		end if
	else

		response.write Sys_IllegalRule1
	end If 

	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then

			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
				response.write "<br>100以上"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
				response.write "<br>80以上未滿100"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
				response.write "<br>60以上未滿80"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
				response.write "<br>40以上未滿60"
			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
				response.write "<br>20以上未滿40"
			else
				response.write "<br>未滿20公里"
			end if

		end If 
		
	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
	end If 

	response.write "</font>"
%></div>
<div id="Layer28" style="position:absolute; left:100px; top:250px; width:217px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:130px; top:280px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" style="position:absolute; left:190px; top:280px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" style="position:absolute; left:270px; top:280px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" style="position:absolute; left:395px; top:290px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write ""
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" style="position:absolute; left:450px; top:330px; width:200px; height:80px; z-index:28"><font size=2><%
	If trim(theBankAccount) <>"" Then
		Response.Write "郵局劃撥戶名："&theBankName&"<br>劃撥帳號："&theBankAccount
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></font></div>

<div id="Layer38" style="position:absolute; left:200px; top:475px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" style="position:absolute; left:360px; top:475px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" style="position:absolute; left:500px; top:475px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>
<div id="Layer41" style="position:absolute; left:670px; top:475px; width:80px; height:12px; z-index:36"><%=Sys_BillFillerMemberID%></div>
</div>
<%
	if (i mod 100)=0 then response.flush
next

If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,4.23,6,4.23,4.23);
</script>