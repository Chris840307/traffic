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
	<tr>
		<td>
		<span style="font-size: 18pt;">Edge安全性設定方式</span>
	</tr>
	<tr bgcolor="#FFFF66">
		<td>
		<span style="font-size: 18pt;">僅初次使用edge瀏覽器時須設定，擇一設定方式即可</span>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td>
		<span style="font-size: 18pt;">設定方式一</span>
		</td>
	</tr>
	<tr>
		<td>
		<span style="font-size: 18pt;">1.開啟Edge後，先進入系統首頁，點選『不安全』，然後將『快顯視窗並自動導向』、『自動下載』設定為允許即可。(如無該選項請參照設定方式二)</span>
		<br/>
		<img src="Image/EdgeSafeSet1.jpg" width="1000">
		</td>
	</tr>
	
	<tr bgcolor="#FFCC33">
		<td>
		<span style="font-size: 18pt;">設定方式二</span>
		</td>
	</tr>
	<tr>
		<td>
		<span style="font-size: 18pt;">1.開啟Edge後，先進入系統首頁，點選『不安全』，然後點選『此網站的權限』進入設定畫面。</span>
		<br/>
		<img src="Image/EdgeSafeSet2.jpg" width="1000">
		</td>
	</tr>
	<tr>
		<td>
		<span style="font-size: 18pt;">2.將下列紅框中的選項全部設定為允許即可。</span>
		<br/>
		<img src="Image/EdgeSafeSet3.jpg" width="1000">
		<br/>
		<img src="Image/EdgeSafeSet4.jpg" width="1000">
		</td>
	</tr>
</table>
