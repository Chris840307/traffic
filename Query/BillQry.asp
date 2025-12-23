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
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

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
	舉發單詳細欄位輸入(匯出 excel檔格式 限制為65536筆) 格式參考99年<br>
	建檔日期<input type="text" value="" name="tDate1">~
	<input type="text" value="" name="tDate2"> 
	<input type="button" value="全部(違規日)" onclick="funchgExecel_YL_all()">
<%
	If sys_City="雲林縣" then
%>
	<input type="button" value="攔停(違規日)" onclick="funchgExecel_YL_s()">
	<input type="button" value="逕舉(違規日)" onclick="funchgExecel_YL_r()">
	<input type="button" value="行人攤販(違規日)" onclick="funchgExecel_YL_p()">
	<input type="button" value="全部(違規日)" onclick="funchgExecel_YL_all()">
<%	else%>
	<input type="button" value="台中市攔停(建檔日)" onclick="funchgExecel_tcc_s()">
	<input type="button" value="台中市逕舉(建檔日)" onclick="funchgExecel_tcc_r()">
	<input type="button" value="台中市行人攤販(建檔日)" onclick="funchgExecel_tcc_p()">
	<input type="button" value="基隆攔停逕舉(建檔日)" onclick="funchgExecel_GL_1()">
	<input type="button" value="基隆行人攤販(建檔日)" onclick="funchgExecel_GL_2()">
	<input type="button" value="基隆裁罰(建檔日)" onclick="funchgExecel_GL_3()">
	<input type="button" value="基隆闖紅燈、超速逕舉(建檔日)" onclick="funchgExecel_GL_4()">
	<input type="button" value="花蓮43(違規日)" onclick="funchgExecel_HL_43()">
	<input type="button" value="花蓮103(違規日)" onclick="funchgExecel_HL_432()">
	<input type="button" value="台南市(違規日)" onclick="funchgExecel_TN3()">
	<br>
<%	End If %>
	<hr>
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	function funchgExecel_tcc_r(){
		UrlStr="BillQry_Execel_Tcc_r.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_tcc_s(){
		UrlStr="BillQry_Execel_Tcc_s.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_tcc_p(){
		UrlStr="BillQry_Execel_Tcc_p.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_YL_r(){
		UrlStr="BillQry_Execel_YL_r.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_YL_all(){
		UrlStr="BillQry_Execel_YL_all.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_YL_s(){
		UrlStr="BillQry_Execel_YL_s.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_YL_p(){
		UrlStr="BillQry_Execel_YL_p.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_GL_1(){
		UrlStr="BillQry_Execel_GL_1.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_GL_2(){
		UrlStr="BillQry_Execel_GL_2.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_GL_3(){
		UrlStr="BillQry_Execel_GL_3.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_GL_4(){
		UrlStr="BillQry_Execel_GL_4.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}
	function funchgExecel_HL_43(){
		UrlStr="BillQry_Execel_HL_43.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}

	function funchgExecel_HL_432(){
		UrlStr="BillQry_Execel_HL_432.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
		newWin(UrlStr,"inputWin",980,550,0,0,"yes","yes","yes","no");
	}

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
	function funchgExecel_TN3(){
		UrlStr="BillQry_Execel_NT3.asp?date1="+myForm.tDate1.value+"&date2="+myForm.tDate2.value;
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
