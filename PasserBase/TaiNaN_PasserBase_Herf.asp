<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
	<%
	sys_unitname=""
	strSQL="select UnitName from Unitinfo where UnitID in(select UnitTypeID from Unitinfo where Unittypeid='"&Session("Unit_ID")&"')"
	set rs=conn.execute(strSQL)
	If not rs.eof Then sys_unitname=trim(rs("UnitName"))
	rs.close

	If instr(sys_unitname,"交通警察大隊")>0 Then
		sys_unitname="交通警察大隊"
	End if
	
	%>
<script language="JavaScript">
location.href="http://10.130.83.148/traffic/PasserBase/PasserBaseQry.asp?Unit_Name=<%=sys_unitname%>&Ch_Name=<%=session("Ch_Name")%>&chk_City=台南市";
</script>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<BODY>

</BODY>
</HTML>
