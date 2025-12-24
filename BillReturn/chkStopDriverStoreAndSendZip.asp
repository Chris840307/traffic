<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If not ifnull(request("Zip")) Then
	
strSQL = "select ZipName from Zip where ZipID='"&trim(request("Zip"))&"'"

	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then%>
		var City='<%=sys_City%>';
		if(myForm.DirverAddress[myForm.chkcnt.value-1].value==''){
			myForm.DirverAddress[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipName"))%>';
		}
		myForm.Sys_ZipName[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipName"))%>';
	<%
	else
		response.write "alert('無此郵地區號請詳查再輸入!!');"
	end if
elseIf not ifnull(request("ZipName")) Then
	
strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(request("ZipName")),5),"臺","台")&"%'"

	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then%>
		var City='<%=sys_City%>';
		myForm.DriverZip[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipID"))%>';
		myForm.Sys_ZipName[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipName"))%>';
		myForm.item[myForm.chkcnt.value].focus();
	<%
	else
		rscnt.close
		
strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(request("ZipName")),3),"臺","台")&"%'"
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then%>
			var City='<%=sys_City%>';
			myForm.DriverZip[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipID"))%>';
			myForm.Sys_ZipName[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipName"))%>';
			alert('該地址無法判別郵遞區號，請確定郵遞區號是否正確!!');
			myForm.item[myForm.chkcnt.value].focus();<%

		else
			response.write "if(myForm.DriverZip[myForm.chkcnt.value-1].value==''){"
			response.write "alert('系統無法判斷郵遞區號，請手動填入來避免資料錯誤!!');}else{myForm.item[myForm.chkcnt.value].focus();}"
		end if
	end if
end if
rscnt.close
conn.close
%>
