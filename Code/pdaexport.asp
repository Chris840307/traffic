<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<HTML>
<HEAD>
<TITLE> 各式資料 匯出 </TITLE>
</HEAD>
<BODY>
<form name=myForm method="post">
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<table width="100%" height="50%" bgcolor="#FFDD77" border="0">
		<tr><td>
			<table width="100%" border="0">
				<tr>
					<td><input type="button" name="btnCar" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 車種代碼" onclick="funCarExport();"></td>
					<td><input type="button" name="btnPolice" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 員警代碼" onclick="funPoliceExport();"></td>
					<td><input type="button" name="btnRule" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 法條代碼" onclick="funRuleExport();"></td>
					<td><input type="button" name="btnStreet" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 違規地點" onclick="funStreetExport();"></td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td><input type="button" name="btnColor" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 顏色代碼" onclick="funColorExport();"></td>
					<td><input type="button" name="btnUnitInfo" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 舉發單位代碼" onclick="funUnitInfoExport();"></td>
					<td><input type="button" name="btnUnitInfo" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 扣件代碼" onclick="funFastenerExport();"></td>
					<td><input type="button" name="btnStation" style="width:160px; height:30px;font-family:標楷體; font-size: 16px; color:#000000;" value="匯出 監理站代碼" onclick="funStationExport();"></td>
					<td>&nbsp;</td>
				</tr>
			</table>
		</td></tr>
	</table>
</font>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funCarExport(){
	UrlStr="BillCarExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPoliceExport(){
	UrlStr="BillPoliceExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funRuleExport(){
	UrlStr="BillRuleExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStreetExport(){
	UrlStr="BillStreetExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funColorExport(){
	UrlStr="BillColorExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funUnitInfoExport(){
	UrlStr="BillUnitInfoExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funFastenerExport(){
	UrlStr="BillFastenerExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funStationExport(){
	UrlStr="BillStationExport_txt.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}
</script>