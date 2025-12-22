<head>
<meta http-equiv="Content-Language" content="zh-tw">

<title>
</title>
</head>
<!--#include virtual="traffic/Common/DB.ini"-->
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	

%>
<table border="1" width="100%" id="table1">
	<tr bgcolor="#FFCC33">
		<td>
		<span style="font-size: 18pt;">Edge相容性設定方式(下列方式擇一即可)</span>
	</tr>
	<tr bgcolor="#CCFFFF">
		<td>
		<span style="font-size: 18pt;">設定方式一</span>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">1.必須先進入系統登入畫面後，點選『在Internet Explorer模式中重新載入』。</span>
		<br/>
		<img src="Image/edgesetB1.jpg" width="1000">
		</td>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">2.將在『相容性檢視中開啟此頁面』以及『下次在Internet Explorer模式中開啟此頁面』都開啟，然後按下『完成』按鈕即可。</span>
		<br/>
		<img src="Image/edgesetB2.jpg" width="1000">
		</td>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">3.設定完成後會有下圖中的工具列，請勿將此工具列關閉。</span>
		<br/>
		<img src="Image/edgesetB3.jpg" width="1000">
		</td>
	</tr>
	<tr bgcolor="#CCFFFF">
		<td>
		<span style="font-size: 18pt;">設定方式二</span>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">1.開啟Edge設定畫面</span>
		<br/>
		<img src="Image/edgeset1.jpg" width="1000">
		</td>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">2.進入預設瀏覽器設定，並請依照下圖做設定，然後點選『新增』按鈕</span>
		<br/>
		<img src="Image/edgeset2.jpg" width="1000">
		</td>
	</tr>
	<tr>
		</td>
		<td>
	<%
	Url1="edgesetA.jpg"
	Url2="edgesetB.jpg"
	Url3="edgesetC.jpg"
	if sys_City="基隆市" then
		Url1="edgeset3.jpg"
		Url2="edgeset4.jpg"
		Url3="edgeset5.jpg"
	elseif sys_City="台南市" then
		Url1="edgeset3TN.jpg"
		Url2="edgeset4TN.jpg"
		Url3="edgeset5TN.jpg"
	end if
	%>
		<span style="font-size: 18pt;"><%
		if sys_City="金門縣" then
		%>3.新增 https://<%=trim(Request.ServerVariables("Local_ADDR") )%>/traffic/ 以及 https://<%=trim(Request.ServerVariables("Local_ADDR") )%>/traffic/traffic_login.asp 兩組網址，新增完成後即可<%
		else%>

		3.新增 http://<%=trim(Request.ServerVariables("Local_ADDR") )%>/traffic/ 以及 http://<%=trim(Request.ServerVariables("Local_ADDR") )%>/traffic/traffic_login.asp 兩組網址，新增完成後即可<%
		end if%></span>
		<br/>
		<img src="Image/<%=Url1%>" width="800">
		<br/>
		<img src="Image/<%=Url2%>" width="800">
		<br/>
		<img src="Image/<%=Url3%>" width="1000">
		</td>
	</tr>
	<tr>
		<td>
		<span style="font-size: 18pt;">4.設定完成後，開啟系統畫面，點選下圖中IE圖示，然後將『在相容性檢視中開啟此頁面』、『下次在Internet Explorer模式中開啟此頁面』選項開啟即可</span>
		<br/>
		<img src="Image/edgeset6.jpg" width="800">
		</td>
	</tr>
	<tr>
		</td>
		<td>
		<span style="font-size: 18pt;">5.設定完成後會有下圖中的工具列，請勿將此工具列關閉。</span>
		<br/>
		<img src="Image/edgesetB3.jpg" width="1000">
		</td>
	</tr>
	
</table>
