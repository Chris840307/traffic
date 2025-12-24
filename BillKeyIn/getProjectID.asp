<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getProjectID.asp
	'專案名稱
	strProject="select Name from Project where ProjectID='"&trim(request("BillProjectID"))&"'"
	set rsProject=conn.execute(strProject)
	if not rsProject.eof then
		ProjectName=trim(rsProject("Name"))
	end if
	rsProject.close
	set rsProject=nothing
%>

setProjectName="<%=ProjectName%>";
if (setProjectName != ""){
	Layer001.innerHTML=setProjectName;
	TDProjectIDErrorLog=0;
}else{
	Layer001.innerHTML=" ";
	TDProjectIDErrorLog=1;
}
<%
conn.close
set conn=nothing
%>
