<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	if sys_City="高雄市" then
		strSQL="select ZipName from Zip where ZipID = "&trim(request("ZipID")) & " and ZipName like '"&trim(sys_City)&"%'"
		set rszip=conn.execute(strSQL)
		if Not rszip.eof Then
	%>
		LayerIllZip.innerHTML="<%=trim(rszip("ZipName"))%>";
		TDIllZipErrorLog=0;
	<%	
		else
	%>
		LayerIllZip.innerHTML=" ";
		TDIllZipErrorLog=1;
	<%
		end if
		rszip.close	
	else
		strSQL="select ZipName from Zip where ZipID = "&trim(request("ZipID"))
		set rszip=conn.execute(strSQL)
		if Not rszip.eof Then
			If left(trim(Request("getIllegalAddress")),3)="台南市" Then 
				strChkIA="select ZipName from Zip where ZipName like '"&left(trim(Request("getIllegalAddress")),5)&"%'"
				Set rsChkIA=conn.execute(strChkIA)
				If Not rsChkIA.eof Then
					getIllegalAddress1=Replace(trim(Request("getIllegalAddress")),Trim(rsChkIA("ZipName")),"")
				Else
					getIllegalAddress1=trim(Request("getIllegalAddress"))
				End If 
				rsChkIA.close
				Set rsChkIA=Nothing 
			Else
				getIllegalAddress1=trim(Request("getIllegalAddress"))
			End If 
			response.write "myForm."&trim(Request("getZipName"))&".value='"&rszip("ZipName")&"'+'"&getIllegalAddress1&"';"
			'response.write "myForm."&trim(Request("getZipName"))&".select();"
		else
			response.write "alert(""郵遞區號錯誤!!"");"
		end if
		rszip.close	
	end if
%>

