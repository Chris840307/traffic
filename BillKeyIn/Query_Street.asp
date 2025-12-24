<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>違規地點代碼查詢</title>
<%
Server.ScriptTimeout=6000
Response.flush
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">違規地點代碼查詢　*可用上、下鍵選擇，enter鍵確定，ESC鍵可重新輸入路段</td>
			</tr>
			<tr>
				<td colspan="4">
				路段代碼<input type="text" name="StreetID" size="6" value="">
				路段名稱<input type="text" name="StreetName" value="">
				<input type="submit" value="查詢" onclick="select_street();">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" value="" name="kinds">
				<br>
				<font size="2" color="clred">可用關鍵字查詢. 代碼開頭碼有大小寫區分。
						例如: 輸入 九如 可查詢所有路段名稱中有 九如 的路段</font>
				</td>
			</tr>
			<tr bgcolor="#FFCC33">
				<td colspan="4">違規地點代碼列表</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="25%" align="center">代碼</td>
				<td width="75%" align="center">路段( 依照名稱排列) </td>
			</tr>
<%
if trim(request("kinds"))="DB_select" then
	strPlus=""
	if trim(request("StreetName"))<>"" then
		strPlus=" where Address Like '%"&trim(request("StreetName"))&"%' "
	end if
	if trim(request("StreetID"))<>"" then
		IF strPlus="" then
			strPlus=" where StreetID Like '"&trim(request("StreetID"))&"%' "
		else
			strPlus=strPlus&" and StreetID Like '"&trim(request("StreetID"))&"%' "
		end if
	end if
	
	If trim(sys_City)="台東縣" Then
		if Session("UnitLevelID")<>"1" then
			IF strPlus="" then
				strPlus=" where UnitID='"&session("Unit_ID")&"'"
			else
				strPlus=strPlus&" and UnitID='"&session("Unit_ID")&"'"
			end if
		end if
	elseif trim(sys_City)="高雄縣" Then
		IF strPlus="" then
			strPlus=" where UnitID='"&session("Unit_ID")&"'"
		else
			strPlus=strPlus&" and UnitID='"&session("Unit_ID")&"'"
		end if
	End if
	strProject="select StreetID,Address from Street "&strPlus&" order by StreetID,Address"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr id="trStreet" title="請點選.." onclick="Inert_Data('<%=trim(rsProject("StreetID"))%>','<%=trim(rsProject("Address"))%>');" <%lightbarstyle 1 %>>
				<td id="tdStreetID" bgcolor="#FFFFCC" align="center"><%=trim(rsProject("StreetID"))%></td>
				<td id="tdAddress"><%=trim(rsProject("Address"))%></td>
			</tr>
<%		Response.flush
	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
elseif trim(request("kinds"))="" and (trim(request("OStreet"))<>"" or trim(request("OStreetID"))<>"") then
	strPlus=""
	if trim(request("OStreet"))<>"" then
		strPlus=" where Address Like '%"&trim(request("OStreet"))&"%' "
	end if
	if trim(request("OStreetID"))<>"" then
		IF strPlus="" then
			strPlus=" where StreetID Like '"&trim(request("OStreetID"))&"%' "
		else
			strPlus=strPlus&" and StreetID Like '"&trim(request("OStreetID"))&"%' "
		end if
	end if
	If trim(sys_City)="台東縣" Then
		if Session("UnitLevelID")<>"1" then
			IF strPlus="" then
				strPlus=" where UnitID='"&session("Unit_ID")&"'"
			else
				strPlus=strPlus&" and UnitID='"&session("Unit_ID")&"'"
			end if
		end if
	elseif trim(sys_City)="高雄縣" Then
		IF strPlus="" then
			strPlus=" where UnitID='"&session("Unit_ID")&"'"
		else
			strPlus=strPlus&" and UnitID='"&session("Unit_ID")&"'"
		end if
	End If
	
	strProject="select StreetID,Address from Street "&strPlus&" order by StreetID,Address"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr id="trStreet" title="請點選.." onclick="Inert_Data('<%=trim(rsProject("StreetID"))%>','<%=trim(rsProject("Address"))%>');" <%lightbarstyle 1 %>>
				<td id="tdStreetID" bgcolor="#FFFFCC" align="center"><%=trim(rsProject("StreetID"))%></td>
				<td id="tdAddress"><%=trim(rsProject("Address"))%></td>
			</tr>
<%		Response.flush
	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing

end if
%>

		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var trStreetIndex=0;
function KeyDown(){
	if (event.keyCode==13){ //Enter換欄
		if(document.all.StreetID.value==''&&document.all.StreetName.value==''){
			event.keyCode=0;
			event.returnValue=false;
			if(document.all.tdStreetID[trStreetIndex]){
				Inert_Data(tdStreetID[trStreetIndex].innerHTML,tdAddress[trStreetIndex].innerHTML);
			}else{
				Inert_Data(tdStreetID.innerHTML,tdAddress.innerHTML);
			}
		}
	}else if (event.keyCode==27){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			document.all.StreetID.focus();
	}else if (event.keyCode==38){ //上換欄
		event.keyCode=0;
		event.returnValue=false;
		document.all.StreetID.blur();
		document.all.StreetName.blur();
		if(trStreetIndex!=0){
			document.all.trStreet[trStreetIndex].style.backgroundColor='#FFFFFF';
			document.all.trStreet[trStreetIndex-1].style.backgroundColor='#CCFFFF';
			trStreetIndex=trStreetIndex-1;
			if(trStreetIndex%14==0){
				event.keyCode=33;
				event.returnValue=true;
			}
		}
	}else if (event.keyCode==40){ //下換欄
		event.keyCode=0;
		event.returnValue=false;
		document.all.StreetID.blur();
		document.all.StreetName.blur();
		if(trStreetIndex<document.all.trStreet.length-1){
			document.all.trStreet[trStreetIndex].style.backgroundColor='#FFFFFF';
			document.all.trStreet[trStreetIndex+1].style.backgroundColor='#CCFFFF';
			trStreetIndex=trStreetIndex+1;
			if(trStreetIndex%14==0){
				event.keyCode=34;
				event.returnValue=true;
			}
		}
	}
}
function select_street(){
	myForm.kinds.value="DB_select";
	myForm.submit();
}
function Inert_Data(SCode,SStreet){
	opener.myForm.IllegalAddress.focus();
	opener.myForm.IllegalAddressID.value=SCode;
	opener.myForm.IllegalAddress.value=SStreet;
	window.close();
}
</script>
</html>
