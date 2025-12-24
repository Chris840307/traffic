<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getVIPCar.asp
	'是否為特殊用車
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	SpecNote=""
	strVIP="select * from SpecCar where CarNo='"&trim(request("CarID"))&"' and RecordStateID<>-1"
	set rsVIP=conn.execute(strVIP)
	if not rsVIP.eof then
		CarCnt=1
		SpecNote=trim(rsVIP("Note"))
	else
		CarCnt=0
	end if
	rsVIP.close
	set rsVIP=nothing
%>setVIPCar("<%=CarCnt%>");
<%
conn.close
set conn=nothing
%>
//是否為特殊用車
function setVIPCar(CarCnt)
{
	if (CarCnt > 0){
		Layer7.innerHTML="＊業管車輛";
		TDVipCarErrorLog=1;
<%if sys_City="雲林縣" then %>
		alert("此車牌為業管車輛!");
		document.myForm.CarNo.select();
<%elseif sys_City="高雄市" then %>
		alert("此車牌為業管車輛。原因 ：<%=SpecNote%>");
		//document.myForm.CarNo.select();
<%end if%>
	}else{
		Layer7.innerHTML=" ";
		TDVipCarErrorLog=0;
	}
}
