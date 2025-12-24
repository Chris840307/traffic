<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
dim cnt

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If not ifnull(request("DriverZip")) Then
	
strSQL = "select ZipName from Zip where ZipID='"&trim(request("DriverZip"))&"'"

	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then%>
		var City='<%=sys_City%>';
		if(myForm.DriverAddress[myForm.chkcnt.value-1].value==''){
			myForm.DriverAddress[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipName"))%>';
		}
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
		<%If sys_City="南投縣" Then%>
			myForm.mailNumber[myForm.chkcnt.value-1].focus();
		<%else%>
			myForm.item[myForm.chkcnt.value].focus();
		<%End if%>
	<%
	else
		rscnt.close
		
strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(request("ZipName")),3),"臺","台")&"%'"
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then%>
			var City='<%=sys_City%>';
			myForm.DriverZip[myForm.chkcnt.value-1].value='<%=trim(rscnt("ZipID"))%>';
			alert('該地址無法判別郵遞區號，請確定郵遞區號是否正確!!');
			<%If sys_City="南投縣" Then%>
				myForm.MailNumber[myForm.chkcnt.value-1].focus();
			<%else%>
				myForm.item[myForm.chkcnt.value].focus();
			<%End if
		else
			response.write "if(myForm.DriverZip[myForm.chkcnt.value-1].value==''){"
			response.write "alert('系統無法判斷郵遞區號，請手動填入來避免資料錯誤!!');myForm.DriverZip[myForm.chkcnt.value-1].focus();}else{myForm.item[myForm.chkcnt.value].focus();}"
		end if
	end if
end if
rscnt.close
%>
var BackIndex=myForm.Sys_BackCause[myForm.chkcnt.value-1].selectedIndex;

var OwnerZip=myForm.OwnerZip[myForm.chkcnt.value-1].value;
var OwnerAddress=myForm.OwnerAddress[myForm.chkcnt.value-1].value;

var DriverZip=myForm.DriverZip[myForm.chkcnt.value-1].value;
var DriverAddress=myForm.DriverAddress[myForm.chkcnt.value-1].value;

if(OwnerAddress!=DriverAddress && DriverAddress!=''){
	myForm.ChkSend[myForm.chkcnt.value-1].value='是';
	
	myForm.Sys_BackCause[myForm.chkcnt.value-1].options[0]=new Option('其他寄存原因','T');
	myForm.Sys_BackCause[myForm.chkcnt.value-1].options[0].selected=true;
	myForm.Sys_BackCause[myForm.chkcnt.value-1].length=1;
}else{
	myForm.ChkSend[myForm.chkcnt.value-1].value='否';
	<%
		strSQL="select ID,Content from DCICode where TypeID=7 and ID in('1','2','3','4','8','M','K','L','O','P','Q')"
		set rs1=conn.execute(strSQL)
		selectCmt=0
		while Not rs1.eof
			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('"&trim(rs1("Content"))&"','"&trim(rs1("ID"))&"');"& vbcrlf
			
			selectCmt=selectCmt+1
			rs1.movenext
		wend
		rs1.close

		if sys_City="南投縣" or sys_City="雲林縣" then
			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('車主死亡','AddReason1');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('拒收','AddReason2');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('公司無此車號','AddReason3');"& vbcrlf

			selectCmt=selectCmt+1

			response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options["&selectCmt&"]=new Option('無此公司','AddReason4');"& vbcrlf

			selectCmt=selectCmt+1
		end if

		response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].length="&selectCmt&";"& vbcrlf
		response.write "myForm.Sys_BackCause[myForm.chkcnt.value-1].options[BackIndex].selected=true;"& vbcrlf
	%>
}
