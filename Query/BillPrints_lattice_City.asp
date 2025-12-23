<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%AuthorityCheck(233)%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI 資料交換紀錄</title>
</head>
<body>
	<object id=factory style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="..\smsx.cab#Version=6,1,432,1"></object>
	<iframe src="" ID="topFrame" width="100%" height="100%" frameborder=0></iframe>
</body>
</html>
<script language="javascript">
	var space=",";
	var strMemID="<%=request("PBillSN")%>";
	var ck_MemID=strMemID.split(space);
	var cnt=0;
	function DP(){
		if(cnt<ck_MemID.length){
			document.getElementById("topFrame").src="BillPrints_lattice_TaiChungCity.asp?PBillSN="+ck_MemID[cnt];
			cnt=cnt+1;
		}
	}
	DP();
</script>