<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>處理交通違規陳情、陳述統計表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body >
<form name=myForm method="post">
<table width="100%" height="100%" border="0" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#FFCC33" colspan="2" height="36">處理交通違規陳情、陳述統計表</td>
	</tr>
	<tr>
		<td bgcolor="#FFEF44">
			統計日期 
		</td>
		<td bgcolor="#FFFFFF" height="86">
			民國 <input type="text" value="" name="sys_Year" size="6"> 年<br><br>
			<input type="radio" name="sys_Date" value="1"> 上半年( 1~6月 ) &nbsp; &nbsp; 
			<input type="radio" name="sys_Date" value="2"> 下半年( 7~12月 )
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFEF44" >
			統計單位
		</td>
	
		<td bgcolor="#FFFFFF" height="35">
			<select name="sys_Unit">
				<option value="">所有單位</option>
<%
	strU1="select * from UnitInfo where showorder in (0,1) order by unitid"
	Set rsU1=conn.execute(strU1)
	If Not rsU1.Bof Then rsU1.MoveFirst 
	While Not rsU1.Eof
%>
				<option value="<%=Trim(rsU1("UnitID"))%>" name="sys_Unit"><%=Trim(rsU1("UnitName"))%></option>
<%
		rsU1.MoveNext
	Wend
	rsU1.close
	set rsU1=nothing
%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFF66" colspan="2" align="center">
			<input type="button" value="產生報表" onclick="fun_report();">
			<input type="button" value="離開" onclick="window.close();">
		</td>
	</tr>
</table>
</form>
</body>
</html>
<script type="text/javascript" src="../js/engine.js"></script>
<script language="javascript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}

function fun_report(){
	var error=0;
	var errorString="";
	var sDate="";
	if(myForm.sys_Year.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入統計年份!!";
	}
	if(myForm.sys_Date(0).checked==false && myForm.sys_Date(1).checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入統計月份!!";
	}
	if (error>0){
		alert(errorString);
	}else{
		if (myForm.sys_Date(0).checked==true){
			sDate="1";
		}else{
			sDate="2";
		}
		UrlStr="ArgueCaseReport_Execel.asp?sys_Year="+myForm.sys_Year.value+"&sys_Date="+sDate+"&sys_Unit="+myForm.sys_Unit.value;
		newWin(UrlStr,"inputWinRep2",900,550,50,10,"yes","yes","yes","no");
	}
}
</script>
<%conn.close%>