<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
	end if 
end if 
rsUInfo.close
set rsUInfo=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"

	If sys_City="台南市" and trim(Sys_UnitID)="07A7" Then
		strSQL="select * from UnitInfo where UnitID='0707'"
	End if
	
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then
	Sys_SendBillSN=request("hd_BillSN")
else
	Sys_SendBillSN=request("BillSN")
End if

if request("DB_Selt")="Save" then

	Sys_SN=Split(request("Sys_SN"),",")
	Sys_ACCEPTDATE=Split(request("Sys_ACCEPTDATE"),",")
	Sys_OPENGOVNUMBER=Split(request("Sys_OPENGOVNUMBER"),",")

	for i=0 to Ubound(Sys_SN)

		strSQL="delete PASSERETAX where billsn="&Sys_SN(i)&" and to_char(ACCEPTDATE,'yyyy')=to_char("&funGetDate(gOutDT(Sys_ACCEPTDATE(i)),0)&",'yyyy')"

		conn.execute(strSQL)


		strIns="insert into PASSERETAX(SN,BillSN,BillNO,OPENGOVNUMBER,ACCEPTDATE,RECORDDATE" &_
			",RECORDMEMBERID)" &_
			" values((select nvl(max(SN),0)+1 from PASSERETAX)," & Sys_SN(i) &_
			",(select billno from passerbase where sn="&Sys_SN(i)&"),'"&trim(Sys_OPENGOVNUMBER(i))&"'"&_
			","&funGetDate(gOutDT(Sys_ACCEPTDATE(i)),0)&",sysdate,"&Session("User_ID")&")" 

		conn.execute(strIns)
	next
	response.write "<script language=""JavaScript"">"
	response.write "alert (""修改完成!!"");"
	Response.Write "self.close();"
	response.write "</script>"
	Response.End
end if


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


%>
<TITLE> 裁決批次套印 </TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">裁決批次套印</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>

						收文日期：
						<input name="def_ACCEPTDATE" type="text" class="btn1" size="10" value=""  onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:25px; height:20px;" onclick="OpenWindow('def_ACCEPTDATE');">
						收文文號：
						<input name="def_OPENGOVNUMBER" type="text" class="btn1" size="20" value="">
						<br><br>
						<input type="button" name="btnSelt" value="整批設定" class="btn3" style="width:100px; height:30px;" onclick="funDefuDate();">
						<input type="button" name="btnOK1" class="btn3" style="width:100px; height:30px;" value="確定存檔" onclick="funSelt();">
					</td>										
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">案件收文列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'><%
				strSQL="select SN,billno,recorddate from passerbase where exists(select 'Y' from passerSend where billsn=passerbase.sn) and exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn)"
				filecnt=0
				set rs=conn.execute(strSQL)
				While not rs.eof
					filecnt=filecnt+1
					Response.Write "<tr><td>"
					Response.Write right("00000"&filecnt,4)
					Response.Write "</td><td>"
					response.write "單號<input name=""Sys_Billno"" class=""btn1"" type=""text"" size=""20"" value="""&trim(rs("billno"))&""" Readonly>"
					
					response.write "<input type=""Hidden"" name=""Sys_SN"" value="""&trim(rs("SN"))&""">"
					Response.Write "</td><td>"
					Response.Write "收文日期："
					Response.Write "<input name=""Sys_ACCEPTDATE"" type=""text"" class=""btn1"" size=""10"" value=""""  onkeyup=""chknumber(this);"">"
					Response.Write "</td><td>"
					Response.Write "收文文號："
					Response.Write "<input name=""Sys_OPENGOVNUMBER"" type=""text"" class=""btn1"" size=""20"" value="""">"

					
					Response.Write "</td></tr>"
					rs.movenext
				Wend
				rs.close
				%>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td height="20" bgcolor="#FFDD77">
		</td>
	</tr>
</table>
<input type="Hidden" name="BillSN" value="<%=Sys_SendBillSN%>">
<input type="Hidden" name="Sys_SN" value="">
<input type="Hidden" name="Sys_Billno" value="">
<input type="Hidden" name="Sys_ACCEPTDATE" value="">
<input type="Hidden" name="Sys_OPENGOVNUMBER" value="">
</form>
<form name="upForm" method="post">

	<input type="Hidden" name="Sys_SN" value="">
	<input type="Hidden" name="Sys_ACCEPTDATE" value="">
	<input type="Hidden" name="Sys_OPENGOVNUMBER" value="">
	<input type="Hidden" name="DB_Selt" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var sys_City="<%=sys_City%>";

function funDefuDate(){
	for(i=0;i<eval("myForm.Sys_ACCEPTDATE").length;i++){
		eval("myForm.Sys_ACCEPTDATE["+i+"]").value=eval("myForm.def_ACCEPTDATE").value;
		eval("myForm.Sys_OPENGOVNUMBER["+i+"]").value=eval("myForm.def_OPENGOVNUMBER").value;
	}
}

function funSelt(){
	var err=0;

	var Sys_SN='';
	var Sys_ACCEPTDATE='';
	var Sys_OPENGOVNUMBER='';

	for(i=0;i<myForm.Sys_ACCEPTDATE.length;i++){
		if(myForm.Sys_OPENGOVNUMBER[i].value!=''){
			if(myForm.Sys_ACCEPTDATE[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行收件日不可空白!!");
				break;
			}
		}
	}
	if(err==0){
		for(i=0;i<myForm.Sys_OPENGOVNUMBER.length;i++){
			if(myForm.Sys_SN[i].value!='' && myForm.Sys_OPENGOVNUMBER[i].value!=''){
				if(Sys_SN!=''){
					Sys_SN=Sys_SN+',';
					Sys_ACCEPTDATE=Sys_ACCEPTDATE+',';
					Sys_OPENGOVNUMBER=Sys_OPENGOVNUMBER+',';
				}

				Sys_SN=Sys_SN + myForm.Sys_SN[i].value;
				Sys_ACCEPTDATE=Sys_ACCEPTDATE + myForm.Sys_ACCEPTDATE[i].value;
				Sys_OPENGOVNUMBER=Sys_OPENGOVNUMBER + myForm.Sys_OPENGOVNUMBER[i].value;
			}
		}

		upForm.Sys_SN.value=Sys_SN;
		upForm.Sys_ACCEPTDATE.value=Sys_ACCEPTDATE;
		upForm.Sys_OPENGOVNUMBER.value=Sys_OPENGOVNUMBER;
		upForm.DB_Selt.value="Save";
		upForm.submit();
	}
}
</script>