<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<%
Server.ScriptTimeout=6000
if trim(request("DB_Add"))="ADD" then
strSQL="update Billbasedcireturn set updatetypeid=rownum where billno in (select billno from (select billno,carno,count(*) cnt from billbasedcireturn where exchangetypeid='W' and billno is not null group by billno,carno) where cnt>1) and exchangetypeid='W' and billno in (select billno from dcilog where batchnumber='"&trim(request("sys_Batchnumber"))&"')"
conn.execute(strSQL)

tmpBillNo="":chkBillNo=""
strSQL="select updatetypeid,Billno from billbasedcireturn where updatetypeid is not null and exchangetypeid='W' and billno in (select billno from (select billno,count(*) cnt from billbasedcireturn where exchangetypeid='W' and billno in (select billno from dcilog where batchnumber='"&trim(request("sys_Batchnumber"))&"') group by billno,Carno) where cnt>1) order by billno"
set rs=conn.execute(strSQL)
While not rs.eof
	If trim(rs("BillNo"))<>trim(tmpBillNo) Then
		tmpBillNo=trim(rs("BillNo"))
	elseif trim(rs("BillNo"))=trim(tmpBillNo) then
		strSQL="delete BillBaseDciReturn where updatetypeid="&trim(rs("updatetypeid"))&" and billno='"&trim(rs("BillNo"))&"'"
		conn.execute(strSQL)
	End if
	rs.movenext
Wend
rs.close
strSQL="update billbasedcireturn set updatetypeid=null where updatetypeid is not null"
conn.execute(strSQL)

strSQL="update Billbasedcireturn set updatetypeid=rownum where billno in (select billno from (select billno,carno,count(*) cnt from billbasedcireturn where exchangetypeid='N' and billno in (select billno from dcilog where batchnumber='"&trim(request("sys_Batchnumber"))&"') group by billno,carno) where cnt>1) and exchangetypeid='N' and billno is not null"
conn.execute(strSQL)

tmpBillNo="":chkBillNo=""
strSQL="select updatetypeid,Billno from billbasedcireturn where updatetypeid is not null and exchangetypeid='N' and billno in (select billno from (select billno,count(*) cnt from billbasedcireturn where exchangetypeid='N' and billno in (select billno from dcilog where batchnumber='"&trim(request("sys_Batchnumber"))&"') group by billno,carno) where cnt>1) order by billno"
set rs=conn.execute(strSQL)
While not rs.eof
	If trim(rs("BillNo"))<>trim(tmpBillNo) Then
		tmpBillNo=trim(rs("BillNo"))
	elseif trim(rs("BillNo"))=trim(tmpBillNo) then
		strSQL="delete BillBaseDciReturn where updatetypeid="&trim(rs("updatetypeid"))&" and billno='"&trim(rs("BillNo"))&"'"
		conn.execute(strSQL)
	End if
	rs.movenext
Wend
rs.close
strSQL="update billbasedcireturn set updatetypeid=null where updatetypeid is not null"
conn.execute(strSQL)
response.Write "處理完成!!"
end if 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>資料修改</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
批號 <input type="text" value="" name="sys_Batchnumber" size="12">
<input type="button" name="save" value="處理" onclick="funAdd();">
<input type="Hidden" name="DB_Add" value="">
</form>
</body>
<script language="javascript">
function funAdd(){
	var err=0;
	if(myForm.sys_Batchnumber.value==''){
		err=1;
		alert("批號不可空白");
	}else{
		myForm.DB_Add.value='ADD';
		myForm.submit();
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}

</script>
</html>