<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>車籍資料列表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="CarDataSelect" then
	strwhere=Session("PrintCarDataSQL")&" and a.CarNo like '%"&trim(request("SelCarNo"))&"%'"
else
	strwhere=Session("PrintCarDataSQL")	
end if
	Session.Contents.Remove("PrintCarDataSQLxls")
	Session("PrintCarDataSQLxls")=strwhere	
	dcitype=trim(request("dcitype"))

'	strdata=" and (substr(e.ownerid,2,1)<>'A' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'S' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'D' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'F' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'G' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'H' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'J' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'K' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'L' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Z' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'X' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'C' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'V' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'B' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'N' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'M' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Q' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'W' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'E' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'R' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'T' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Y' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'U' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'I' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'O' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'P' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>' ' "
'	strdata=strdata&")"

'	strdata2=" and (substr(e.ownerid,1,1)='A' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='S' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='D' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='F' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='G' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='H' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='J' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='K' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='L' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Z' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='X' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='C' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='V' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='B' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='N' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='M' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Q' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='W' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='E' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='R' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='T' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Y' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='U' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='I' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='O' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='P' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)=' ' "
'	strdata2=strdata2&")"
	strdata2=strdata2&"  and (substr(e.ownerid,2,1) in ('1','2','3','4','5','6','7','8','9','0'))"
	strdata2=strdata2&" and (substr(e.ownerid,1,1) in ('A','S','D','F','G','H','J','K','L','Z','X','C','V','B','N','M','Q','W','E','R','T','Y','U','I','O','P',' '))"

	strSQL="select distinct e.billno,e.ownerid,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='') "&strwhere&" order by a.RecordDate"

	set rsfound=conn.execute(strSQL)

	strCnt="select count(*) as cnt from (select distinct a.SN,a.BillTypeID,a.CarSimpleID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,a.IllegalDate,a.imagefilenameb,a.Note,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='')"&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere
'response.write strSQL
%>

</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
	<tr>
		<td bgcolor="#FFCC33">
			<font size="3"><strong>車籍資料清冊</strong></font>
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
			筆 (共 <%=DBsum%> 筆)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			車號：<input type="text" size="10" name="SelCarNo" value="<%=trim(request("SelCarNo"))%>" onkeyup="this.value=this.value.toLocaleUpperCase()">
			<input type="button" name="Sel1" value="查詢" onclick="CarDataSelect();">
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="850" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>單號</th>
					<th>車號</th>
					<th>證號</th>
					<th>車主姓名</th>
					<th>車主地址</th>

				</tr>
				<tr bgcolor="#FFFFFF" align="center" >
				<%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					ListSN=DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						ListSN=ListSN+1
						if rsfound.eof then exit for
%>						<tr bgcolor="#ffffff" <%lightbarstyle 0 %>>

							<td width="20"><%=rsfound("BillNo")%></td>
							<td width="70"><%=rsfound("CarNo")%></td>
							<td width="50"><%=rsfound("OwnerID")%></td>
							<td width="120"><%=funcCheckFont(rsfound("Owner"),20,1)%></td>
							<td width="400"><%
							if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
								response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),20,1)
							end if
							%></td>
						</tr>
<%						Response.flush
						rsfound.movenext
					next
				%>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFDD77" align="center">
			<a href="file:///.."></a>
			<a href="file:///......"></a>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<img src="space.gif" width="13" height="8">

			<img src="space.gif" width="5" height="8">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="button" name="btnExecel" value=" 離 開 " onclick="window.close();">
		</td>
	</tr>

</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="hidden" name="strSQL" value="<%=tmpSQL%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	function funChkVIP(){
		window.open("DciCarDataChkSpecCar.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
	}
function funChkVIP_HL(){
	window.open("DciCarDataChkSpecCar_HL.asp","chk_vip1","width=620,height=440,left=200,top=150,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
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
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function funchgExecel(){
	UrlStr="PrintOwnerDataList_Excel.asp";
	newWin(UrlStr,"PrintWin_xls",790,550,50,10,"yes","yes","yes","no");
}
function funchgList(){
	UrlStr="PrintCarDataList_HL_A3.asp";
	newWin(UrlStr,"DciPrintWin_xls",790,550,50,10,"yes","yes","yes","no","yes");
}
function CarDataSelect(){
	if (myForm.SelCarNo.value==""){
		alert("請輸入車號！");
	}else{
		myForm.DB_Move.value=0;
		myForm.kinds.value="CarDataSelect";
		myForm.submit();	
	}
}
function DelBill_NoDCI(DelSN){
	UrlStr="BillBase_Del.asp?DBillSN="+DelSN;
	newWin(UrlStr,"DelBillBasePage1",290,100,200,110,"no","no","yes","no","no");
}
</script>
<%conn.close%>