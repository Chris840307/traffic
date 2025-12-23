<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單管理</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%


%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">
	<b>審計室資料匯出</b>
	<br>
	<br>
	舉發單祥細欄位輸入(匯出 excel檔格式 限制為65536筆) 格式參考99年<br>
	建檔日期<input type="text" value="" name="tDate1">~
	<input type="text" value="" name="tDate2"> 
	<input type="button" value="台南+金門監理站(建檔日)" onclick="funchgExecel_tcc()">
	__
	<input type="button" value="屏東結案(填單日)" onclick="funchgExecel()">
	__
	<input type="button" value="屏東結案_違規日(審計)" onclick="funchgExecel_PD2()">
	__
	<input type="button" value="逕舉 舉發明細_建檔日" onclick="funchgExecelbilltype()"> 
	__
	<input type="button" value="攔停 舉發明細_違規日" onclick="funchgExecel3()"> 
	__
	<input type="button" value="逕舉 舉發明細_違規日" onclick="funchgExecel4()"> 
	__
	<input type="button" value="舉發明細_填單日(南投)" onclick="funchgExecel_TN()"> 
	__
	<input type="button" value="逕舉 舉發明細_違規日(台南縣)" onclick="funchgExece2_TN()"> 
	__
	<input type="button" value="99舉發明細_違規日(彰化)" onclick="funchgExecel_ch()"> 
	__
	<input type="button" value="逕舉 舉發明細_違規日(審計)" onclick="funchgExece2_ch()"> 
	<br>
	<hr>
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

	function funchgExecel(){
		UrlStr="BillQry_Execel_YL.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_PD(){
		UrlStr="BillQry_Execel_PD.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_PD2(){
		UrlStr="BillQry_Execel_PD2.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecelbilltype(){
		UrlStr="BillQry2_Execel.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel3(){
		UrlStr="BillQry3_Execel.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel4(){
		UrlStr="BillQry4_Execel.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_TN(){
		UrlStr="BillQry_Execel_NT1.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExece2_TN(){
		UrlStr="BillQry_Execel_NT2.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_ch(){
		UrlStr="BillQry_Execel_CH99.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExece2_ch(){
		UrlStr="BillQry_Execel_CH2.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	
	function funchgExecel_tcc(){
		UrlStr="BillQry_Execel_TCC.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
	}
</script>
