<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 10px; color:#ff0000;}
.style2 {font-family:"標楷體"; font-size: 10px}
.style3 {font-family:"標楷體"; font-size: 14px}
.style4 {font-family:"標楷體"; font-size: 18px}
.style7 {font-family:"標楷體"; font-size: 13px}
.style8 {font-family:"標楷體"; font-size: 36px}
.style11 {font-family:"標楷體"; font-size: 14px}
.style15 {font-family:"標楷體"; font-size: 15px}
-->
</style>
</head>

<body>
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

sys_title="宜蘭縣政府警察局"

for i=0 to Ubound(PBillSN)
if cint(i)<>0 then response.write "<div class=""PageNext"">　</div>"
	
	strBil="select distinct BillSN,BillNo,CarNo,BatchNumber from PasserDcilog where BillSN="&PBillSN(i)&" and ExchangetypeID='W' and BillTypeID=2 and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1)"

set rsbil=conn.execute(strBil)

If rsbil.eof Then 
	rsbil.close
	Response.Write "<span class=""style8"">案件還沒上傳監理站入案。</span>"
	Response.End
end If 

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


strSql="select a.LoginID,c.Content,a.ChName,b.UnitLevelID,b.UnitID,b.UnitName,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b,(select ID,Content from Code where TypeID=4 ) c where a.UnitID=b.UnitID and a.JobID=c.ID(+) and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillJobName=trim(mem("Content"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	SysUnit=unit("UnitName")
	SysUnitTel=trim(unit("Tel"))
	SysUnitAddress=trim(unit("Address"))
end if
unit.close


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

imgfile="":arrfile=""
Sys_IisImagePath="":Sys_ImageFileNameA="":Sys_ImageFileNameB=""
strSQL="select IisImagePath,ImageFileNameA ImageFileName1,ImageFileNameB ImageFileName2,ImageFileNameC ImageFileName3 from PasserIllegalImage where billsn="&trim(rsbil("BillSN"))
set rsimage=conn.execute(strSQL)
if Not rsimage.eof then
	Sys_IisImagePath=trim(rsimage("IisImagePath"))
	For k=1 To 3
		If Trim(rsimage("ImageFileName"&k))<>"" Then

			If imgfile<>"" Then imgfile=imgfile&"@"
			imgfile=imgfile&rsimage("ImageFileName"&k)
		End if
	Next

	If imgfile<>"" Then

		arrfile=Split(imgfile,"@")

		Sys_ImageFileNameA=trim(arrfile(0))

		If InStr(imgfile,"@") > 0 Then Sys_ImageFileNameB=trim(arrfile(1))

	End if
end If 

rsimage.close

Sys_MailNumber=0

'strSQL="select min(SendMailStation) SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rsbil("BillSN"))
'
'set rsSend=conn.execute(strSQL)
'
'if Not rsSend.eof then
'	Sys_MailNumber=rsSend("SendMailStation")
'
'end If 
'rsSend.close

DelphiASPObj.GenBillPrintBarCode1 PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber&"95000017","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate

rsbil.close

firstBacrCode=right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&"D"&Sys_StationID

pageleft=0
pagetop=0
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->

<div id="L78" class="pageprint" style="position:relative;">
<div id="Layer01" class="style3" style="position:absolute; left:<%=75+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write sys_title&SysUnit
%>
</div>
<div id="Layer66" class="style3" style="position:absolute; left:<%=345+pageleft%>px; top:<%=0+pagetop%>px; z-index:10"><%
	response.Write SysUnitAddress&"(郵戳請勿蓋在條碼上)"
%>
</div>
<div id="Layer000" style="position:absolute; left:30px; top:0px; z-index:1"><%

	strURL="il01_legal_city_People.jpg"

	Response.Write "<img src=""..\legal_Img\"&strURL&""" width=""715"" height=""1290"">"
	
	%>
</div>
<div id="Layer01" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=15+pagetop%>px; z-index:10"><%
	response.write funcCheckFont(Sys_Owner,16,1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_CarNo&"<br>"
	response.write Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
%>
</div>

<div id="Layer02" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=60+pagetop%>px; z-index:11"><%
	response.write Sys_BillNo%>
</div>

<div id="Layer03" class="style3" style="position:absolute; left:<%=180+pageleft%>px; top:<%=50+pagetop%>px; z-index:10"><%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"" width=""180"" height=""25"">"%>
</div>
<!--
<div id="Layer05" class="style3" style="position:absolute; left:<%=470+pageleft%>px; top:<%=287+pagetop%>px; z-index:10">
　　<b>第<%=Sys_MailNumber%>號</b><br>
　　<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>><br>
　<b><%="　"&Sys_MAILCHKNUMBER%></b>
</div>
-->
<div id="Layer03" class="style3" style="position:absolute; left:<%=405+pageleft%>px; top:<%=270+pagetop%>px; z-index:10"><%
	Response.Write "<img src=""../image/cutter.jpg""></img>"%>
</div>

<!---------------------------------- 第一段送達證書到這邊------------------------------------->

<div id="Layer04" class="style3" style="position:absolute; left:<%=110+pageleft%>px; top:<%=290+pagetop%>px; z-index:10"><%
	response.write funcCheckFont(Sys_Owner,16,1)
	Response.Write "　台啟&nbsp&nbsp&nbsp&nbsp"
	Response.Write "<span class=""style4""><B>郵遞區號："
	response.write Sys_OwnerZip
	Response.Write "</B></span>"
	%>
</div>

<div id="Layer05" class="style3" style="position:absolute; left:<%=340+pageleft%>px; top:<%=330+pagetop%>px; z-index:10"><%
	Response.Write "<img  src=""..\BarCodeImage\"&Sys_BillNo&"_1.jpg"">"%>
</div>

<div id="Layer06" class="style3" style="position:absolute; left:<%=100+pageleft%>px; top:<%=310+pagetop%>px; width:340px; z-index:10"><%
	'--------------------------------------如果是抓 戶籍補正的資料-----------------------------------------------------------------------------------------------------------
			response.write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
			response.write "<br><br>"
	%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:<%=110+pageleft%>px; top:<%=330+pagetop%>px; z-index:10"><%
	Response.Write sys_title&SysUnit
%>
</div>

<div id="Layer07" class="style3" style="position:absolute; left:<%=70+pageleft%>px; top:<%=360+pagetop%>px; z-index:10"><%
	'Response.Write Sys_CarNo&"　"&Sys_A_Name&"　"&Sys_CarColor&"　"&Sys_STATIONNAME
	Response.Write Sys_CarNo&"　"&Sys_STATIONNAME
%>
</div>

<div id="Layer08" class="style3" style="position:absolute; left:<%=380+pageleft%>px; top:<%=360+pagetop%>px; z-index:10"><%
	Response.Write Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)%>
</div>

<%if trim(Sys_ImageFileNameA)<>"" then%>
	<div id="Layer09" style="position:absolute; left:38px; top:480px; z-index:5"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameA&""" width=""390"" height=""280"">"
	%></DIV>
<%end if%>

<%if trim(Sys_ImageFileNameB)<>"" then%>
	<div id="Layer10" style="position:absolute; left:430px; top:480px; z-index:1"><%
		response.write "<img src="""&Sys_IisImagePath&Sys_ImageFileNameB&""" width=""300"" height=""238"">"
	%></DIV>
<%end if%>

<div id="Layer2" style="position:absolute; left:<%=50+pageleft%>px; top:<%=845+pagetop%>px; width:202px; height:36px; z-index:10">Ｖ</div>

<div id="Layer4" style="position:absolute; left:<%=170+pageleft%>px; top:<%=830+pagetop%>px; width:202px; height:36px; z-index:10">v</div>

<div id="Layer9" class="style3" style="position:absolute; left:<%=45+pageleft%>px; top:<%=865+pagetop%>px; z-index:10"><%
	response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"" width=""200"" height=""35"">"
%></div>

<div id="Layer26" class="style3" style="position:absolute; left:<%=115+pageleft%>px; top:<%=897+pagetop%>px; z-index:10"><%

	response.write firstBacrCode

%></div>

<div id="Layer10" style="position:absolute; left:<%=510+pageleft%>px; top:<%=845+pagetop%>px; z-index:10"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"" width=""200"" height=""30"""%>></div>

<div id="Layer11" class="style3" style="position:absolute; left:<%=540+pageleft%>px; top:<%=890+pagetop%>px; z-index:10"><%="宜警交"&"　　　　"&Sys_BillNo%></div>

<div id="Layer12" class="style7" style="position:absolute; left:<%=110+pageleft%>px; top:<%=920+pagetop%>px; width:150px; height:11px; z-index:10"><%
		response.write "逕行舉發&nbsp;"
%>
</div>

<div id="Layer13" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=915+pagetop%>px; width:28px; height:11px; z-index:10"><%'=Sys_Sex%></div>
<div id="Layer14" class="style3" style="position:absolute; left:<%=370+pageleft%>px; top:<%=915+pagetop%>px; width:324px; height:10px; z-index:10"><%'="*本單可至郵局或委託代收之超商繳納"%></div>

<div id="Layer15" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=925+pagetop%>px; width:100px; height:10px; z-index:10"><%'if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></div>
<div id="Layer16" class="style3" style="position:absolute; left:<%=425+pageleft%>px; top:<%=925+pagetop%>px; width:106px; height:13px; z-index:10"><%'=Sys_DriverID%></div>
<div id="Layer17" class="style3" style="position:absolute; left:<%=620+pageleft%>px; top:<%=925+pagetop%>px; width:99px; height:12px; z-index:10"><%'=fastring%></div>
<div id="Layer18" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=965+pagetop%>px; width:100px; height:14px; z-index:10"><%=Sys_CarNo%></div>
<div id="Layer19" class="style3" style="position:absolute; left:<%=260+pageleft%>px; top:<%=965+pagetop%>px; width:117px; height:20px; z-index:10"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style3" style="position:absolute; left:<%=500+pageleft%>px; top:<%=965+pagetop%>px; width:300px; height:17px; z-index:10"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" class="style3" style="position:absolute; left:<%=125+pageleft%>px; top:<%=990+pagetop%>px; width:507px; height:13px; z-index:10"><%=Sys_OwnerZip&" "& funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)%></div>

<div id="Layer22" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:13px; z-index:10"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style3" style="position:absolute; left:<%=170+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:17px; z-index:10"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style3" style="position:absolute; left:<%=220+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:16px; z-index:10"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style3" style="position:absolute; left:<%=270+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:16px; z-index:10"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" class="style3" style="position:absolute; left:<%=320+pageleft%>px; top:<%=1010+pagetop%>px; width:40px; height:13px; z-index:10"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style3" style="position:absolute; left:<%=400+pageleft%>px; top:<%=1015+pagetop%>px; width:340px; height:31px; z-index:10"><%
	
	if (trim(Sys_Rule1)="72000011" or trim(Sys_Rule1)="72000021" or trim(Sys_Rule1)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"
		end if
	else
		
		response.write Sys_IllegalRule1
	end if


	if (trim(Sys_Rule2)="72000011" or trim(Sys_Rule2)="72000021" or trim(Sys_Rule2)="72000031") then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then

			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里"

		end If 
		
	elseif trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then

		response.write "<br>"&Sys_IllegalRule2
	end If 
%></div>
<div id="Layer28" class="style3" style="position:absolute; left:<%=115+pageleft%>px; top:<%=1030+pagetop%>px; width:220px; height:15px; z-index:10"><span class="style3"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" class="style3" style="position:absolute; left:<%=120+pageleft%>px; top:<%=1065+pagetop%>px; width:50px; height:11px; z-index:10"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style3" style="position:absolute; left:<%=210+pageleft%>px; top:<%=1065+pagetop%>px; width:35px; height:13px; z-index:10"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style3" style="position:absolute; left:<%=280+pageleft%>px; top:<%=1065+pagetop%>px; width:32px; height:15px; z-index:10"><%=Sys_DealLineDate(2)%></div>

<div id="Layer32" class="style3" style="position:absolute; left:<%=405+pageleft%>px; top:<%=1085+pagetop%>px; z-index:10"><%
	response.write left(trim(Sys_Rule1),2)&"　&nbsp;"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if

%></div>

<div id="Layer33" class="style7" style="position:absolute; left:<%=440+pageleft%>px; top:<%=1135+pagetop%>px; z-index:11"><%
	If trim(theBankAccount) <>"" Then
		Response.Write "1.郵局匯票戶名："&theBankName
		Response.Write "<br>2."&Sys_STATIONNAME&"臨櫃繳納。"
	else
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_STATIONNAME&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write Sys_StationTel
	End if 
%></div>

<div id="Layer35" style="position:absolute; left:<%=400+pageleft%>px; top:<%=1175+pagetop%>px; width:100px; height:49px; z-index:10"><%
		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"" nowrap>宜蘭縣政府警察局<br>"&SysUnit&"</td></tr>"
		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" class=""style1"">TEL"&SysUnitTel&"</td></tr>"
		response.write "</table>"
	%></div>
<div id="Layer36" style="position:absolute; left:<%=610+pageleft%>px; top:<%=1210+pagetop%>px; width:100px; height:43px; z-index:10"><%
'	if instr(Sys_BillNo,"QZ")>0 then
'			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>小隊長&nbsp;林添福</span></td></tr>"
'			response.write "</table>"
'	elseif Sys_UnitID="TO00" then
'			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>組長&nbsp;莊松杰</span></td></tr>"
'			response.write "</table>" 

	'elseif trim(Session("Unit_ID"))="TN00" then
			
		
	'elseif trim(Sys_UnitLevelID)="1" then
'	else

'		response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
'		response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">交通違規專用章<br>"&Sys_jobName&"&nbsp;"&Sys_MainChName&"</span></td></tr>"
'		response.write "</table>"
'	end if
%></div>
<div id="Layer37" style="position:absolute; left:<%=610+pageleft%>px; top:<%=1190+pagetop%>px; width:200px; height:46px; z-index:10"><%
		if instr(Sys_BillNo,"QZ")>0 then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">警員&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		elseif  trim(Session("Unit_ID"))="TG01" then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">警務佐&nbsp;梁建泰</span></td></tr>"
			response.write "</table>"

		elseif trim(Session("Unit_ID"))="TP00" then
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		else
			response.write "<table style=""border-color:#ff0000;border-style:solid;"" border=""1"" cellspacing=0 cellpadding=0>"
			response.write "<tr><td style=""border-color:#ff0000;border-style:solid;"" width=100 height=25 align=""center""><span class=""style1"">"&Sys_BillJobName&"&nbsp;"&Sys_ChName&"</span></td></tr>"
			response.write "</table>"

		end if%></div>
<div id="Layer38" class="style3" style="position:absolute; left:<%=210+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:10px; z-index:10"><%=sys_Date(0)%></div>
<div id="Layer39" class="style3" style="position:absolute; left:<%=365+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:13px; z-index:10"><%=sys_Date(1)%></div>
<div id="Layer40" class="style3" style="position:absolute; left:<%=515+pageleft%>px; top:<%=1255+pagetop%>px; width:60px; height:11px; z-index:10"><%=sys_Date(2)%></div>
<div id="Layer41" class="style3" style="position:absolute; left:<%=690+pageleft%>px; top:<%=1255+pagetop%>px; width:80px; height:12px; z-index:10"><%=Sys_BillFillerMemberID%></div>
<div id="Layer43" class="style3" style="position:absolute; left:<%=300+pageleft%>px; top:<%=1285+pagetop%>px; width:250px; height:12px; z-index:10"><%=Sys_DCIRETURNCARTYPE%></div>
</div>

</div>

<%
	if (i mod 10)=0 then response.flush
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
	//window.print();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>