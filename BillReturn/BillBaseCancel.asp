<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>整批資料徹銷送達</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
-->
</style>
</HEAD>
<BODY>
<%
Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
if sys_City="花蓮縣" then
	CarName="姓名"
else
	CarName="車號"
end if

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(Ucase(request("item")),",")
	sys_batchnumber=""
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" then
			If ifnull(sys_batchnumber) Then
				strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
				set rsSN=conn.execute(strSN)
				if not rsSN.eof then
					sys_batchnumber=(year(now)-1911)&"N"&trim(rsSN("SN"))
				end if
				rsSN.close
				set rsSN=nothing
			End if
			
			strSQL="select SN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate from billbase where billno='"&trim(Sys_BillNo(i))&"' and recordstateid=0"

			set rs=conn.execute(strSQL)

			RecordDate=funGetDate(rs("RecordDate"),1)

			strInsStoG="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,ExchangeDate,ExchangeTypeID,ReturnMarkType,DCIwindowName,BatchNumber) values(DCILOG_SEQ.nextval,"&trim(rs("SN"))&",'"&trim(rs("BillNo"))&"',"&trim(rs("BillTypeID"))&",'"&trim(rs("CarNo"))&"','"&trim(rs("BillUnitID"))&"',"&RecordDate&","&Session("User_ID")&",sysdate,'N','Y','"&Session("DCIwindowName")&"','"&sys_batchnumber&"')" 

			conn.execute strInsStoG

			strUpdStoG="update billmailhistory set MailTypeID=null where BillSN="&trim(rs("SN"))

			conn.execute strUpdStoG

			rs.close
		end if
	next
	Response.write "<script>"
	Response.Write "alert('徹銷批號："&sys_batchnumber&"\n儲存完成！');"
	Response.write "</script>"
end if
%>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>整批資料徹銷送達</strong></td>
	</tr>

	<tr>
		<td height="26" bgcolor="#FFCC33">徹銷紀錄列表 <b> <%=titleStr%></b></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:400px;background:#FFFFFF">
				<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap></td>
					</tr>
				</table>
			</Div>
		</td>
	</tr>
	<tr align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="btnOK1" value="確定存檔" onclick="funSelt();">
			<input type="button" name="insert2" value="再多30筆" onClick="insertRow(fmyTable)">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;
function insertRow(isTable){
	for(i=0;i<=29;i++){
		Rindex = isTable.rows.length;
		if(isTable.rows.length>0){
			Cindex = isTable.rows[Rindex-1].cells.length;
		}else{
			Cindex=0;
		}
		if(Rindex==0||Cindex==1){
			nextRow = isTable.insertRow(Rindex);
			txtArea = nextRow.insertCell(0);
		}else{
			if(cunt==0){
				Cindex=0;
				isTable.rows[Rindex-1].deleteCell();
			}
			txtArea =isTable.rows[Rindex-1].insertCell(Cindex);
		}
		cunt++;
		//txt_nameStr = "item"+cunt;
		var cnt_num=("0000"+cunt).substr(("0000"+cunt).length-3,3);
		txtArea.innerHTML =cnt_num+".&nbsp;單號<input type=text name='item' size=10 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=10 class='btn1' readOnly><br><br>";
	}
}

function keyFunction(itemcnt) {
	//myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkBillBaseCancel.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}




function funSelt(){
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}

for(j=0;j<=3;j++){
	insertRow(fmyTable);
}
</script>