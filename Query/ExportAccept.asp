<!--#include virtual="traffic/Common/SqlServerdb.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_攔停點收匯出檔.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=big5">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 14">
<link rel=File-List href="file3457.files/filelist.xml">
<style id="活頁簿1_17813_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font517813
	{color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;}
.xl1517813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6417813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt hairline windowtext;
	border-bottom:.5pt hairline windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6517813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt hairline windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6617813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt hairline windowtext;
	border-bottom:.5pt hairline windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6717813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt hairline windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6817813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6917813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt hairline windowtext;
	border-right:.5pt hairline windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7017813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt hairline windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7117813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7217813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt hairline windowtext;
	border-right:.5pt hairline windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7317813
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt hairline windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt hairline windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-char-type:none;}
-->
</style>
</head>

<body>
<!--[if !excel]>　　<![endif]-->
<!--下列資訊是由 Microsoft Excel 網頁發佈精靈所產生。-->
<!--如果由 Excel 重新發佈相同的項目時，在 DIV 標籤間的所有資訊將會被取代。-=
->
<!----------------------------->
<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->
<!----------------------------->

<div id="活頁簿1_17813" align=center x:publishsource="Excel">

<table border=0 cellpadding=0 cellspacing=0 width=738 style='border-collapse: collapse;table-layout:fixed;width:554pt'>
 <col width=72 span=3 style='width:54pt'>
 <col width=90 style='mso-width-source:userset;mso-width-alt:2880;width:68pt'>
 <col width=72 span=6 style='width:54pt'>
 <tr class=xl6817813 height=34 style='mso-height-source:userset;height:25.5pt'>
  <td height=34 class=xl6417813 width=72 style='height:25.5pt;width:54pt'>序</td>
  <td class=xl6517813 width=72 style='border-left:none;width:54pt'>違規時間</td>  
  <td class=xl6617813 width=72 style='width:54pt'>違規單號</td>
  <td class=xl6517813 width=72 style='border-left:none;width:54pt'>違規人證號</td>
  <td class=xl6617813 width=90 style='border-left:none;width:68pt'>車號/引擎號</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>車種</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>砂石註記</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>違規法條</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>違規地點</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>扣件內容</td>
  <td class=xl6617813 width=72 style='border-left:none;width:54pt'>舉發人</td>
  <td class=xl6717813 width=72 style='border-left:none;width:54pt'>送達狀態</td>
 </tr>
 <%
 strSQL="select (vl_date+vl_time) illegaldate,vl_bil_no billno,vl_id_idn DriverID,plte CarNo" & _
	",(vl_ord1+(case when vl_ord2 <>'' then ','+vl_ord2 else '' end)+(case when vl_ord3 <>'' then ','+vl_ord3 else '' end)) rule1" & _
	",(vl_town+vl_plce) IllegalAddress" & _
	",(pd_arm_no1+(case when pd_arm_no2 <>'' then ','+pd_arm_no2 else '' end)+(case when pd_arm_no3 <>'' then ','+pd_arm_no3 else '' end)+(case when pd_arm_no4 <>'' then ','+pd_arm_no4 else '' end)) BillMem" & _
 " from BOOKING_TICKETS" & _
 " where vl_bil_no is not null and convert(datetime,created_time) between '" & gOutDT(request("RecordDate1")) & " 00:00:00' and '" & gOutDT(request("RecordDate2")) & " 23:59:59'" & _
 " order by created_time"

 set rs=conn.execute(strSQL)
 cnt=0
 While not rs.eof
	cnt=cnt+1
	Response.Write "<tr height=34 style='height:25.5pt'>"
	Response.Write "<td height=34 class=xl6917813 align=right width=72 style='height:25.5pt; border-top:none; width:54pt'>"
	Response.Write cnt
	Response.Write "</td>"

	Response.Write "<td class=xl7017813 width=72 style='border-top:none;border-left:none; width:54pt'>"
	Response.Write rs("illegaldate")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("billno")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("DriverID")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("CarNo")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("rule1")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("IllegalAddress")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write rs("BillMem")
	Response.Write "</td>"

	Response.Write "<td class=xl7117813 style='border-top:none'>"
	Response.Write "</td>"

	Response.Write "</tr>"

	response.flush
	rs.movenext 
 Wend
 rs.close
%>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=90 style='width:68pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
  <td width=72 style='width:54pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>
