<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

if request("DB_Selt")="Save" then
	BillSN=Split(request("BillSN"),",")
	theJudeDate=gOutDT(request("Sys_JudeDate"))
	strSQL="Update UnitInfo set WordNum='"&trim(request("Sys_WordNum"))&"' where UnitID='"&Session("Unit_ID")&"'"
	conn.execute(strSQL)

	'strSQL="Update UnitInfo set ManageMemberName='"&trim(request("Sys_UnitChName"))&"' where UnitID='"&DB_UnitID&"'"
	'conn.execute(strSQL)

	session("Sys_UnitChName")=request("Sys_UnitChName")

	for i=0 to Ubound(BillSN)
		strSQL="select Sn,BillNo from PasserBase where SN="&BillSN(i)
		set rs=conn.execute(strSQL)
		theJudeNo=""
		theJudeForFeit=0
		strSQL="Select OpenGovNumber,ForFeit from PasserJude where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
		set rsJude=conn.execute(strSQL)
		if trim(request("Sys_JudeNo"))="2" then
			theJudeNo=rs("BillNo")
		end if
		if Not rsJude.eof then
			if trim(rsJude("ForFeit"))<>"" then theJudeForFeit=rsJude("ForFeit")
		end if
		rsJude.close

		strSQL="Select * from PasserUrge where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
		set rsRuge=conn.execute(strSQL)
		if rsRuge.eof then
			strIns="insert into PasserUrge(BillSN,BillNO,ForFeit,OpenGovNumber,UrgeDate" &_
				",BIGUNITBOSSNAME,RecordStateID,RecordDate,RecordMemberID)" &_
				" values("&trim(rs("Sn"))&",'"&trim(rs("BillNo"))&"',"&theJudeForFeit&_
				",'"&trim(theJudeNo)&"',TO_DATE('"&theJudeDate&"','YYYY/MM/DD')"&_
				",'"&trim(request("Sys_UnitChName"))&"',0,sysdate,'"&Session("User_ID")&"')"
				conn.execute(strIns)
		else
			if trim(request("Sys_JudeNo"))<>"" then
				strUpd="update PasserUrge set OpenGovNumber='"&trim(theJudeNo)&"' where BillSN="&trim(rs("Sn"))&" and BillNo='"&trim(rs("BillNo"))&"'"
			end if
			conn.execute(strUpd)

			strSQL="Update PasserUrge set UrgeDate=TO_DATE('"&theJudeDate&"','YYYY/MM/DD') where BillSN="&trim(rs("Sn"))&" and BillNo='"&trim(rs("BillNo"))&"' and UrgeDate is null"
			conn.execute(strSQL)
		end if
		rsRuge.close
		rs.close
	next
	response.write "<script language=""JavaScript"">"
	response.write "window.opener.funJudeList();"
	response.write "</script>"
	Response.End
end if
If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
else
	Sys_SendBillSN=request("hd_BillSN")
End if
SysWordNum=""
strSQL="select WordNum from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If Not rs.eof Then SysWordNum=trim(rs("WordNum"))
rs.close
%>
<TITLE> 催告批次套印 </TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#1BF5FF" class="pagetitle">催告批次套印</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<font color="Red"><B>交裁字號：</B></font>
						<input name="Sys_WordNum" type="text" class="btn1" size="10" maxlength="15" value="<%=SysWordNum%>">交字第
						<br>
						文號產生規則<input class="btn1" type="radio" name="Sys_JudeNo" value="1"<%if trim(request("Sys_JudeNo"))="1" then response.write " checked"%>>
						不用文號：
						<input class="btn1" type="radio" name="Sys_JudeNo" value="2"<%if trim(request("Sys_JudeNo"))="2" or trim(request("Sys_JudeNo"))="" then response.write " checked"%>>
						舉發單為文號
						<br>
						承辦人&nbsp;
						<input name="Sys_ChName" type="text" class="btn1" size="10" maxlength="12" value="<%
							if trim(request("Sys_Chmem"))<>"" then
								response.write trim(request("Sys_ChName"))
							else
								response.write trim(Session("Ch_Name"))
							end if
						%>">
						單位主管&nbsp;
						<input name="Sys_UnitChName" type="text" class="btn1" size="10" maxlength="12" value="<%
							if trim(request("Sys_UnitChName"))<>"" then
								response.write trim(request("Sys_UnitChName"))
							else
								strSQL="select ManageMemberName,SecondManagerName,UnitName from UnitInfo where UnitID='"&DB_UnitID&"'"
								set rsUnit=conn.execute(strSQL)
								sHelpUnitName=rsUnit("UnitName")
								if Not rsUnit.eof then response.write rsUnit("SecondManagerName")
								rsUnit.close
							end if
						%>">
						催繳日期&nbsp;<input name="Sys_JudeDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%
							if trim(request("Sys_JudeDate"))<>"" then
								response.write gInitDT(trim(request("Sys_JudeDate")))
							else
								response.write gInitDT(date)
							end if
						%>">
						應到案處所&nbsp;
						<select name="Sys_DutyUnit" class="btn1">
							<option value="">請選取</option>
							<%strSQL="select UnitID,UnitName from UnitInfo"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("UnitID")&""""
								if isEmpty(request("Sys_DutyUnit")) then
									if trim(Session("Unit_ID"))=trim(rs1("UnitID")) then response.write " selected"
								else
									if trim(request("Sys_DutyUnit"))=trim(rs1("UnitID")) then response.write " selected"
								end if
								response.write ">"&rs1("UnitName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
					</td>
				</tr>
				<tr>
					<td>
						<input class="btn1" type="checkbox" name="Sys_PasserNotify" value="1">
						交辦單
						<input class="btn1" type="checkbox" name="Sys_PasserSign" value="1">
						簽辦單
						<input class="btn1" type="checkbox" name="Sys_PasserUrge" value="1">
						催繳通知書
						<input class="btn1" type="checkbox" name="Sys_PasserDeliver" value="1">
						送達證書
						<input class="btn1" type="checkbox" name="Sys_PasserSend" value="1">
						寄存通知
						<input class="btn1" type="checkbox" name="Sys_PasserLabel_miaoli" value="1">
						保防標籤
						<input type="button" name="btnSelt" value="確定" onclick="funSelt();">
						<input name="Submit433222" type="button" class="style3" value=" 關 閉 " onclick="self.close();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td height="20" bgcolor="#1BF5FF">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="BillSN" value="<%=Sys_SendBillSN%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
function funSelt(){
	if(myForm.BillSN.value!=''){
		if(myForm.Sys_PasserNotify.checked){
			opener.myForm.Sys_PasserNotify.value="1";
		}else{
			opener.myForm.Sys_PasserNotify.value="";
		}
		if(myForm.Sys_PasserUrge.checked){
			opener.myForm.Sys_PasserUrge.value="1";
		}else{
			opener.myForm.Sys_PasserUrge.value="";
		}
		if(myForm.Sys_PasserDeliver.checked){
			opener.myForm.Sys_PasserDeliver.value="1";
		}else{
			opener.myForm.Sys_PasserDeliver.value="";
		}
		if(myForm.Sys_PasserSend.checked){
			opener.myForm.Sys_PasserSend.value="1";
		}else{
			opener.myForm.Sys_PasserSend.value="";
		}
		if(myForm.Sys_PasserSign.checked){
			opener.myForm.Sys_PasserSign.value="1";
		}else{
			opener.myForm.Sys_PasserSign.value="";
		}
		opener.myForm.Session_JudeName.value=myForm.Sys_ChName.value;
		opener.myForm.BillUrge.value="1";
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
</script>