<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/bannernodata.asp"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單績效檢核</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style2 {font-size: 18px}
.style3 {font-size: 15px}
.style5 {
	font-size: 14px;
	font-weight: bold;
	color: #FF0000;
}
.style9 {font-family: "標楷體"}
.style10 {font-family: "標楷體"; font-size: 15px; }
.style9 {font-family: "標楷體"; font-weight: bold;}
-->
</style></head>
<body onload="ctlUnit();">
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

str_DayID="RecordDate,IllegalDate,DeallIneDate,BillFillDate,ExChangeDate,MailDate,MailAcceptDate,MailReturnDate,StoreAndSendMailReturnDate,OpenGovMailReturnDate"
str_DayName="建檔日期,違規日期,應到案日期,填單日期,入案日期,郵寄日期,收受日期,單退日期,寄存日期,公示日期"
tmp_DayID=split(str_DayID,",")
tmp_DayName=split(str_DayName,",")
sql = "select UnitName from UnitInfo where UnitID= '" & Session("Unit_ID") & "'"
Set RSSystem = Conn.Execute(sql)
if Not RSSystem.Eof Then printUnit = RSSystem("UnitName")
RSSystem.close

sql = "Select * From UserRptInfo Where UserId=" & Session("User_ID") & " And UPPER(ReportId)='REPORT0010'"
Set RSSystem = Conn.Execute(sql)
While Not RSSystem.Eof
	if trim(RSSystem("FieldName"))="startDate_q" then startDate_q=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="endDate_q" then endDate_q=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="UnitID_q" then UnitID_q=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="IllegalDate_start" then IllegalDate_start=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="IllegalDate_end" then IllegalDate_end=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="ListOrder" then ListOrder=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="unit" then unit=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="rptHead1" then rptHead1=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="rptHead2" then rptHead2=trim(RSSystem("FieldValue"))
	if trim(RSSystem("FieldName"))="sumDate_q" then sumDate_q=trim(RSSystem("FieldValue"))
	RSSystem.MoveNext
Wend
if trim(startDate_q)="" and trim(endDate_q)="" then startDate_q=tmp_DayID(0):endDate_q=tmp_DayID(1)
RSSystem.close
%>
<Form name="QryBase0004" >
<table width="100%" height="100%" border="0">
	<tr>
		<td height="27" bgcolor="#FFCC33">舉發單績效檢核</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
		<tr>
		<td colspan="3">
			列印時間 : <%=gInitDT(date)%><br>
			列印單位 : <%=printUnit%><br>
			列印人員 : <%=Session("Ch_Name")%><br>
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td colspan="2">
			表頭一
			<input type="text" name="rptHead1" value="<%=rptHead1%>">
			<br>
			表頭二
			<input name="rptHead2" type="text" value="<%=rptHead2%>">
		</td>
	</tr>
	<tr>
		<td colspan="2"><hr></td>
	</tr>
	<tr>
		<td>案件類型<font size="2">&nbsp; ( &nbsp;1~69 條&nbsp; )	</font>&nbsp;&nbsp;
			<input name="BillBaseType" type="radio" value="1" checked>
			欄停案件
			<input name="BillBaseType" type="radio" value="2">
			逕舉案件
			<input name="BillBaseType" type="radio" value="0">
			欄停&nbsp;+&nbsp;逕舉
        </td>	
	</tr>	
<tr>
		<td>案件類型<font size="2"> ( 69條之後 )	</font>&nbsp;&nbsp;			
			<input name="BillBaseType" type="radio" value="9">
			所有案件			
        </td>	
		<td>
			<input name="unit" type="checkbox" value="y" onClick="ctlUnit();"<%if trim(unit)="y" then response.write " checked"%>>&nbsp;
			單位&nbsp;<%
				strSQL="Select UnitID,UnitName from UnitInfo order by UnitID,UnitTypeID"
				set rsUnit=conn.execute(strSQL)
				response.write "<select name=""UnitID_q"">"
				while Not rsUnit.eof
					response.write "<option value='"&trim(rsUnit("UnitID"))&"'"
					if trim(UnitID_q)=trim(rsUnit("UnitID")) then response.write " selected"
					response.write ">"&trim(rsUnit("UnitName"))&"</option>"
					rsUnit.movenext
				wend
				rsUnit.close
			%>
		</td>
	</tr>		
	<tr>
		<td width="34%">統計期間
			<select name="startDate_q" onchange="funDateChante();"><%
				tmpDateName=""
				for i=0 to ubound(tmp_DayID)
					if trim(tmp_DayID(i))<>trim(endDate_q) then
						response.write "<option value="""&tmp_DayID(i)&""""
						if trim(tmp_DayID(i))=trim(startDate_q) then
							response.write " selected"
							tmpDateName=tmp_DayName(i)
						end if
						response.write ">"&tmp_DayName(i)&"</option>"
					end if
				next%>
			</select>

			<input type="Hidden" name="startDate_Name" value="<%=tmpDateName%>">

			<select name="endDate_q" onchange="funDateChante();"><%
				tmpDateName=""
				for i=0 to ubound(tmp_DayID)
					if trim(tmp_DayID(i))<>trim(startDate_q) then
						response.write "<option value="""&tmp_DayID(i)&""""
						if trim(tmp_DayID(i))=trim(endDate_q) then
							response.write " selected"
							tmpDateName=tmp_DayName(i)
						end if
						response.write ">"&tmp_DayName(i)&"</option>"
					end if
				next%>
			</select>
			<input type="Hidden" name="endDate_Name" value="<%=tmpDateName%>">
			&nbsp;&nbsp;&nbsp;
		
			<input name="overtype" type="radio" value="1" checked>
			大於
			<input name="overtype" type="radio" value="2">
			小於	
        			
			<input type='text' size='3' name='sumDate_q' value='<%=sumDate_q%>' maxLength='8'>天
		</td>
		<td width="25%">
		<%if sys_City="南投縣" or sys_City="雲林縣" then%>
			<input type="checkbox" name="MemchkBox" value="y" onClick="ctlMember();" >&nbsp;建檔人代碼
			<input type="text" value="" Name="RecordMemID" size="8" disabled>
		<%end if%>
		</td>
	</tr>

	<tr>
		<td>違規日期
			<input type='text' size='9' name='IllegalDate_start' value='<%=IllegalDate_start%>' maxLength='8'>
			<input name="datestra" type="button" value="..." onclick="OpenWindow('IllegalDate_start');">
			~
			<input type='text' size='9' name='IllegalDate_end' value='<%=IllegalDate_end%>' maxLength='8'>
			<input name="datestrb" type="button" value="..." onclick="OpenWindow('IllegalDate_end');">
		</td>
		<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			排列
			<%
			response.write "<Select Name=""ListOrder"">"
			response.write "<option value=""IllegalDate"""
			if ListOrder="IllegalDate" then response.write " selected"
			response.write ">舉發日期</option>"
			response.write "<option value=""Billmem1"""
			if ListOrder="Billmem1" then response.write " selected"
			response.write ">舉發人</option>"
			response.write "</select>"
			%>
		</td>
	</tr>
      <tr>
      	
        <td colspan="3"><br><font color="red">  <b>逕舉相片</b> 稽核請注意 : 逕舉相片固定相片收取下來後給建檔人員 時通常 違規時間 到 填單日期 / 入案日期 通常會有一段時間 。<b>
        					 <br>請注意 統計期間 </b>稽核的條件選項 </font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center">
        <input type="button" name="Submit42" value="產出報表(輸出格式 Excel )" onclick="chkSend();">
        <input type="button" name="Submit4" value="回到前一頁" onClick="javascript:history.back();">
         </p>    </td>
  </tr>
</Form>  
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language=javascript>
function funDateChante(){
	var space=",";
	var strDateID='<%=str_DayID%>';
	var arr_DateID=strDateID.split(space);
	var strDateName='<%=str_DayName%>';
	var arr_DateName=strDateName.split(space);
	var tmp_startDate_q=QryBase0004.startDate_q.value;
	var tmp_endDate_q=QryBase0004.endDate_q.value;

	QryBase0004.endDate_q.length=0;

	for(var i=0;i<arr_DateID.length;i++){
		if(arr_DateID[i]!=tmp_startDate_q){
			QryBase0004.endDate_q.length=QryBase0004.endDate_q.length+1;
			QryBase0004.endDate_q.options[QryBase0004.endDate_q.length-1]=new Option(arr_DateName[i],arr_DateID[i]);
			if(QryBase0004.endDate_q.options[QryBase0004.endDate_q.length-1].value==tmp_endDate_q){
				QryBase0004.endDate_q.options[QryBase0004.endDate_q.length-1].selected=true;
				QryBase0004.endDate_Name.value=QryBase0004.endDate_q.options[QryBase0004.endDate_q.selectedIndex].text;
			}
		}
	}
	QryBase0004.startDate_q.length=0;
	for(var i=0;i<arr_DateID.length;i++){
		if(arr_DateID[i]!=tmp_endDate_q){
			QryBase0004.startDate_q.length=QryBase0004.startDate_q.length+1;
			QryBase0004.startDate_q.options[QryBase0004.startDate_q.length-1]=new Option(arr_DateName[i],arr_DateID[i]);
			if(QryBase0004.startDate_q.options[QryBase0004.startDate_q.length-1].value==tmp_startDate_q){
				QryBase0004.startDate_q.options[QryBase0004.startDate_q.length-1].selected=true;
				QryBase0004.startDate_Name.value=QryBase0004.startDate_q.options[QryBase0004.startDate_q.selectedIndex].text;
			}
		}
	}
}
function ctlUnit(){
	if (document.all.unit.checked==true){
		document.all.unit.value="y";
		document.all.UnitID_q.disabled=false;
	}else{
		document.all.unit.value="n";
		document.all.UnitID_q.disabled=true;
	}
}

<%if sys_City="南投縣" or sys_City="雲林縣" then%>
function ctlMember(){
	if (document.all.MemchkBox.checked==true){
		//document.all.RecordMemID.value="";
		document.all.RecordMemID.disabled=false;
	}else{
		document.all.RecordMemID.value="";
		document.all.RecordMemID.disabled=true;
	}
}
<%end if%>
function chkSend(){
	var sDate = document.all.startDate_q.value;
	var eDate = document.all.endDate_q.value;	
	var rptHead1 = document.all.rptHead1.value;
	var rptHead2 = document.all.rptHead2.value;
	var startDate_q = document.all.startDate_q.value;
	var endDate_q = document.all.endDate_q.value;
	var unit = document.all.unit.value;
	var UnitID_q = document.all.UnitID_q.value;
	var ReportName = document.all.rptHead2.value;
	var IllegalDate_start = document.all.IllegalDate_start.value;
	var IllegalDate_end = document.all.IllegalDate_end.value;
	var ListOrder= document.all.ListOrder.value;
	var sumDate_q= document.all.sumDate_q.value;
	var chkOk = true;

  if (document.all.rptHead2.value==''){  	  
  	  alert ("您尚未填入【表頭二】！！");
  	  chkOk = false ;
  /*}else if (document.all.IllegalDate_start.value==''){  	  
  	  alert ("您尚未設定【違規日期】範圍！！");
  	  chkOk = false ;
  }else if (document.all.IllegalDate_end.value==''){  	  
  	  alert ("您尚未設定【違規日期】範圍！！");
  	  chkOk = false ;
  }else if (IllegalDate_start > IllegalDate_end){
  		alert('違規日期之起始日期不得大於結束日期');
  		chkOk = false ;*/
  }else if(document.all.startDate_q.value==''){  	  
		alert ("您尚未設定【統計期間】範圍！！");
		chkOk = false ;
  }else if (document.all.endDate_q.value==''){
  		alert ("您尚未設定【統計期間】範圍！！");
		chkOk = false ;
  }else if(document.all.sumDate_q.value==''){
		alert ("您尚未設定超過天數！！");
		chkOk = false ;
  }

  if (chkOk==true) {
    //window.open("","tmpForm","scrollbars=yes,menubar=no,width=1,height=1,resizable=no,left=0,top=0,status=no");
	QryBase0004.action="ResportSave0004.asp";
	QryBase0004.target="tmpForm";
	QryBase0004.submit();
	QryBase0004.action="";
	QryBase0004.target="";
  }
}
</script>