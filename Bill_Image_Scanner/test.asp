<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!-- #include file="../Common/Bannernodata.asp"-->

<form name="myForm">

建檔單位
						<%=SelectUnitOption("Sys_RecordUnit","Sys_RecordMemberID")%>
						<img src="space.gif" width="8" height="10">
						建檔人
						<%=SelectMemberOption("Sys_RecordUnit","Sys_RecordMemberID")%>
</form>

<script type="text/javascript" src="../js/date.js"></script>  

<script language="javascript">
				<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&request("Sys_RecordMemberID")&"');"%>
		<%response.write "UnitMan('Sys_RecordUnit','Sys_RecordMemberID','"&request("Sys_RecordMemberID")&"');"%>
function UnitMan(UnitName,UnitMem,tmpMemID){
	if(document.all[UnitName].value!=''){
		runServerScript("UnitAdd.asp?UnitID="+document.all[UnitName].value+"&UnitMem="+UnitMem+"&MemberID="+tmpMemID);
	}else{
		document.document.all[UnitMem].options[1]=new Option('所有人','');
		document.document.all[UnitMem].length=1;
	}
}
</script>