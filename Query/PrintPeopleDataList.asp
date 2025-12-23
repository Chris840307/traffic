<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>建檔清冊</title>
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>建檔清冊</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>建檔日期</td>
					<td nowrap>
						<input name="RecordDate1" class="btn1" type="text" value="<%=gInitDT(date)%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate1');">
					~
						<input name="RecordDate2" class="btn1" type="text" value="<%=gInitDT(date)%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('RecordDate2');">
					</td>
				</tr>
				<tr>
					<td>到案處所</td>
					<td nowrap>
						<select name="Sys_MemberStation" class="btn1"><%
							strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
							set rsUnit=conn.execute(strSQL)
							If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
							rsUnit.close
							strUnitID="":strtmp=""
							if trim(Session("UnitLevelID"))="1" then
								strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
								strtmp=strtmp+"<option value="""">所有單位</option>"
							else
								strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID=(select UnitTypeid from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"

								set rs1=conn.execute(strSQL)
								while Not rs1.eof
									if trim(strUnitID)="" and rs1("UnitLevelID")>1 then
										strUnitID=trim(rs1("UnitID"))
									elseif rs1("UnitLevelID")>1 then 
										strUnitID=strUnitID&"','"&trim(rs1("UnitID"))
									end if
									rs1.movenext
								wend
								rs1.close
								strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
							end if
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
								if trim(rs1("UnitID"))=trim(request("Sys_MemberStation")) then
									strtmp=strtmp+" selected"
								end if
								strtmp=strtmp+">"&rs1("UnitName")&"</option>"
								rs1.movenext
							wend
							rs1.close
							strtmp=strtmp+"</select>"
							response.write strtmp%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
	<input type="Hidden" name="Record_SQL" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funAdd(){
	var err=0;
	if(myForm.RecordDate1.value==""){
		err=1;
		alert("請輸入建檔日!!");
	}
	if(err==0){
		if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				err=1;
				alert("建檔日輸入不正確!!");
			}
		}
	}
	if(err==0){
		if(myForm.RecordDate2.value!=""){
			if(!dateCheck(myForm.RecordDate2.value)){
				err=1;
				alert("建檔日輸入不正確!!");
			}
		}
	}
	if(err==0){
		myForm.Record_SQL.value="true";
		UrlStr="PrintPeopleDataList_Stop.asp"
		myForm.action=UrlStr;
		myForm.target="DataList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		window.close();
	}
}
</script>