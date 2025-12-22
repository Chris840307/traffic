<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%

if request("DB_Selt")="Selt" then
	
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴項目分析表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body onkeydown="KeyDown()">
<form name="myForm" method="post">
<table width="100%" height="100%" border="0" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#FFCC33" colspan="2" height="36">申訴項目分析表</td>
	</tr>
	<tr>
		<td bgcolor="#FFEF44">
			單位 
		</td>
		<td bgcolor="#FFFFFF" >
			<select name="Sys_Unit" class="btn1">
						<option value="">所有單位</option>
		<%
				strUnit="select * from UnitInfo where ShowOrder in (0,1) order by showorder,Unitid"
				set rsUnit=conn.execute(strUnit)
				If Not rsUnit.Bof Then rsUnit.MoveFirst 
				While Not rsUnit.Eof
		%>
						<option value="<%=trim(rsUnit("UnitID"))%>" <%if trim(request("Sys_Unit"))=trim(rsUnit("UnitID")) then response.write "selected"%>><%=trim(rsUnit("UnitName"))%></option>
		<%
				rsUnit.MoveNext
				Wend
				rsUnit.close
				set rsUnit=nothing
		%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFEF44">
			統計日期 
		</td>
		<td bgcolor="#FFFFFF" >
			<input type="text" name="Date1">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Date1');"> 至
			<input type="text" name="Date2">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Date2');">
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFEF44">
			舉發單類別 
		</td>
		<td bgcolor="#FFFFFF" >
			<input type="radio" name="BillType" value="0" checked>全部
			<br>
			<input type="radio" name="BillType" value="1">攔停
			<br>
			<input type="radio" name="BillType" value="2">逕舉
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFF99" colspan="2" height="36" align="center">
			<input type="button" value="產生報表" onclick="funAdd();">
			<input type="button" value="離    開" onclick="funExt();">
		</td>
	</tr>
</table>

</form>
</body>
</html>
<script type="text/javascript" src="../js/engine.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funAdd(){
	var err=0;
	var tmpBillType=0;
	if(myForm.Date1.value=='' || myForm.Date2.value==''){
		err=1;
		alert("統計日期不可空白!!");
	}
	if(err==0){
		
		if (myForm.BillType(0).checked==true){
			tmpBillType="0";
		}else if(myForm.BillType(1).checked==true){ 
			tmpBillType="1";
		}else{
			tmpBillType="2";
		}
		UrlStr="ArgueCaseAnalyze_Execel.asp?Date1="+myForm.Date1.value+"&Date2="+myForm.Date2.value+"&BillType="+tmpBillType+"&BillUnit="+myForm.Sys_Unit.value;
		newWin(UrlStr,"inputWin4",900,550,50,10,"yes","yes","yes","no");
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,"otherwin","width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}

function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}
</script>
<%conn.close%>