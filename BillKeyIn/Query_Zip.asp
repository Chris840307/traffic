<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>郵遞區號查詢</title>
<%
LawOrder=trim(request("LawOrder"))
theRuleVer=trim(request("RuleVer"))
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
	<form name="myForm" method="post" onsubmit="return DB_Select();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#EBE5FF">
				<td colspan="7">郵遞區號查詢 　*可用上、下鍵選擇，enter鍵確定，ESC鍵可重新輸入法條</td>
			</tr>
			<tr>
				<td colspan="7">
				縣市 <input type="text" size="10" maxlength="10" name="CityID" value="<%=trim(request("CityID"))%>">
				<input type="submit" name="BB1" value="查詢" onclick="DB_Select();">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<br>
				<font size="2" color="clred">可用關鍵字查詢. 。
				</td>
			<tr>
			<tr bgcolor="#EBE5FF">
				<td colspan="7">郵遞區</td>
			</tr>
			<tr bgcolor="#FAFAF5">
				<td width="10%" align="center">代碼</td>
				<td width="9%" align="center">區</td>
			</tr>
<%
if trim(request("kinds"))="DB_Select" or (not ifnull(Request("ZipCity"))) or (not ifnull(Request("ZipCity"))) then
	strZip="select ZipID,ZipName from Zip where 1=1"
	If not ifnull(Request("ZipCity")) Then
		strZip=strZip&" and ZipName like '"&Trim(Request("ZipCity"))&"%'"
	elseIf not ifnull(Request("CityID")) Then
		strZip=strZip&" and ZipName like '"&Trim(Request("CityID"))&"%'"
	End if

	If not ifnull(Request("ZipCity")) Then
		'strZip=strZip&" and ZipID like '%"&Trim(Request("IllegalZip"))&"%'"
	End if
	
	strZip=strZip&" order by ZipID"
	set rsZip=conn.execute(strZip)
	While Not rsZip.Eof
%>
		<tr id="trLaw" title="請點選.." <%
		If Trim(Request("ZipCity"))="高雄市" Then%>
		onclick="Inert_Data2('<%=trim(rsZip("ZipID"))%>','<%=trim(rsZip("ZipName"))%>','<%=trim(request("ObjName"))%>');"
		<%else%>
		onclick="Inert_Data('<%=trim(rsZip("ZipID"))%>','<%=trim(request("ObjName"))%>');"
		<%End if		
		%>  <%lightbarstyle 1 %>>
			<td id="ItemID" bgcolor="#EBE5FF" align="center"><%=trim(rsZip("ZipID"))%></td>
			<td><%=trim(rsZip("ZipName"))%>&nbsp;</td>
		</tr>
<%		rsZip.MoveNext
	Wend
	rsZip.close
	set rsZip=nothing

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
var trLawIndex=0;
function KeyDown(){
	if (event.keyCode==13){ //Enter換欄
		if(document.all.LawID.value==''){
			event.keyCode=0;
			event.returnValue=false;
			Inert_Data(ItemID[trLawIndex].innerHTML,IllegalRule[trLawIndex].innerHTML,Level1[trLawIndex].innerHTML);
		}
	}else if (event.keyCode==27){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			document.all.LawID.focus();
	}else if (event.keyCode==38){ //上換欄
		event.keyCode=0;
		event.returnValue=false;
		document.all.LawID.blur();
		if(trLawIndex!=0){
			document.all.trLaw[trLawIndex].style.backgroundColor='#FFFFFF';
			document.all.trLaw[trLawIndex-1].style.backgroundColor='#CCFFFF';
			trLawIndex=trLawIndex-1;
			if(trLawIndex%11==0){
				event.keyCode=33;
				event.returnValue=true;
			}
		}
	}else if (event.keyCode==40){ //下換欄
		event.keyCode=0;
		event.returnValue=false;
		document.all.LawID.blur();
		if(trLawIndex<document.all.trLaw.length-1){
			document.all.trLaw[trLawIndex].style.backgroundColor='#FFFFFF';
			document.all.trLaw[trLawIndex+1].style.backgroundColor='#CCFFFF';
			trLawIndex=trLawIndex+1;
			if(trLawIndex%11==0){
				event.keyCode=34;
				event.returnValue=true;
			}
		}
	}
}
function DB_Select(){
	myForm.kinds.value="DB_Select";
	myForm.submit();
}
function Inert_Data(ZipID,ObjName){
	eval("opener.myForm."+ObjName).focus();
	eval("opener.myForm."+ObjName).value=ZipID;
	opener.CodeEnter(ObjName);
	window.close();
}
function Inert_Data2(ZipID,ZipName,ObjName){
	eval("opener.myForm."+ObjName).focus();
	eval("opener.myForm."+ObjName).value=ZipID;
	opener.LayerIllZip.innerHTML=ZipName;
	opener.TDIllZipErrorLog=0;
	window.close();
}
myForm.CityID.focus();

</script>
</html>
