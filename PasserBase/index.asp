<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\bannernodata.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>報表主頁</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">


</head>
<%

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If
if request("dbkind")="Del" then
	strSQL = "Delete From UserRptInfo Where UserId=" & request("memID") & " And UPPER(ReportId)='"&Ucase(trim(request("reportID")))&"' and ReportName is null"
	conn.execute(strSQL)
	strSQL = "Delete From UserRptInfo Where UserId=" & request("memID") & " And UPPER(ReportId)='"&Ucase(trim(request("reportID")))&"' and ReportName='"&request("reportName")&"'"
	conn.execute(strSQL)
	strSQL="Delete from UserListInfo where  UserId=" & request("memID") & " And UPPER(ReportId)='"&Ucase(trim(request("reportID")))&"' and ReportName='"&request("reportName")&"'"
	conn.execute(strSQL)
	strSQL="Delete from UserLawInfo where  UserId=" & request("memID") & " And UPPER(ReportId)='"&Ucase(trim(request("reportID")))&"' and ReportName='"&request("reportName")&"'"
	conn.execute(strSQL)
end if

ReportID=split("REPORT0008,REPORT0011,REPORT0003,REPORT0009,REPORT0012,REPORT0013,REPORT0015,REPORTBASE0010,REPORTBASE0011,REPORTBASE0014,REPORTBASE0015",",")

ReportUrl=split("ReportBase0001,ReportBase0005,ReportBase0002,ReportBase0003,ReportBase0006,ReportBase0007,ReportBase0009,ReportBase0010,ReportBase0011,ReportBase0014,REPORTBASE0015",",")
%>
<body>
<form name="myForm">
<table width="100%" border="0">
  <tr>
    <td height="44" bgcolor="#FFCC33"><span class="pagetitle">報表主頁 </span></td>
  </tr>
  <tr>
    <td height="385" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellspacing="1">
	<tr>
		<td height="44" bgcolor="#FFCC33">1.固定報表    <img src="../Image/space.gif" width="20"> 
    <a href="report.doc"><font size="5">-->報表內容過長,如何列印在同一頁 說明文件</font></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0022.asp"><span class="pagetitle">各單位舉發件數總計表 (各單位建檔統計) </span></a></img></td>
  </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0024.asp"><span class="pagetitle">告發件數統計表(單位分局加派出所)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0006.asp"><span class="pagetitle">告發張數統計表(法條別)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0010.asp"><span class="pagetitle">告發張數統計表(員警別明細)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0011.asp"><span class="pagetitle">告發張數統計表(員警別總計)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0001.asp"><span class="pagetitle">違規法條別告發件數統計表(單位法條別)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0013.asp"><span class="pagetitle">違規法條最低罰鍰統計表(單位別)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0012.asp"><span class="pagetitle">違規法條別告發件數統計表(車種別)</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0007.asp"><span class="pagetitle">員警開單件數排行</span></a></img></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0002.asp"><span class="pagetitle">登錄件數統計表（建檔人）入案</span></a></img></td>
	</tr>
	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0033.asp"><span class="pagetitle">登錄件數統計表（建檔人）註記</span></a></img></td>
	</tr>

	<!--<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0005.asp"><span class="pagetitle">單位舉發件數統計表</span></a></img></td>
	</tr>-->
	

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="index_3.asp"><span class="pagetitle">重點工作統計表 與 淨牌績效統計表</span></a></img>(署版)</td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0026.asp" target="_blank"><span class="pagetitle">處理違反道路交通管理事件統計表</span></a></img>(署版)</td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0009.asp" target="_blank"><span class="pagetitle">舉發違反道路交通管理事件成果</span></a></img>(署版)<img src="../Image/space.gif" width="20"></img><a href="npa-1.jpg" target="blank">操作說明</href></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0021.asp" target="_blank"><span class="pagetitle">舉發違反道路交通管理事件成果(97年版)</span></a></img>(署版)<img src="../Image/space.gif" width="20"></img><a href="npa-1.jpg" target="blank">操作說明</href></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0014.asp" target="_blank"><span class="pagetitle">舉發違反高速公路及快速公路管制規定成果表</span></a></img>(署版)<img src="../Image/space.gif" width="20"></img><a href="npa-1.jpg" target="blank">操作說明</href></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0004.asp"><span class="pagetitle">舉發單績效檢核</span></a></img></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0011.asp"><span class="pagetitle">專案統計成果表</span></a></img></td>
    </tr> 

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0015.asp"><span class="pagetitle">中華電信交換數據費用表</span></a></img></td>
    </tr> 

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="Report0016.asp"><span class="pagetitle">上傳監理站次數統計表</span></a></img></td>
    </tr>


	<tr>
		<td height="44" bgcolor="#FFCC33">2.彈性報表<img src="../Image/space.gif" width="20"> 
    <a href="report.doc"><font size="5">-->報表內容過長,如何列印在同一頁 說明文件</font></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0001.asp"><span class="pagetitle">法條路段交叉表（法條在左，支援路段模糊查詢）</span></a></img>　<a href="./SampleImage/ReportBase0001.jpg" target="_blank">範例</a></td>
	</tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10" ><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0005.asp"><span class="pagetitle">法條路段交叉表（法條在上）</span></a></img>　<a href="./SampleImage/ReportBase0005.jpg" target="_blank">範例</a></td>
	</tr>

    <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0003.asp"><span class="pagetitle">法條車種交叉表</span></a></img>　<a href="./SampleImage/ReportBase0003.jpg" target="_blank">範例</a></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0006.asp"><span class="pagetitle">路段車種交叉表</span></a></img></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0002.asp"><span class="pagetitle">法條單位交叉表</span></a></img>　<a href="./SampleImage/ReportBase0002.jpg" target="_blank">範例</a></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0014.asp"><span class="pagetitle">單位法條交叉表</span></a></img>　<a href="./SampleImage/ReportBase0002.jpg" target="_blank">範例</a></td>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0007.asp"><span class="pagetitle">法條人員交叉表</span></a></img>
    </tr>

	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0009.asp"><span class="pagetitle">法條年齡交叉表(分本縣市和非本縣市)</span></a></img>
    </tr>
	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0011.asp"><span class="pagetitle">單位人員專案統計表</span></a><a href="單位人員專案統計表.doc"><span class="pagetitle">(使用說明)</span></a></img>
    </tr>

    <% if sys_City="台南市" then %>
	<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0015.asp"><span class="pagetitle">單位人員專案統計表(攔停)</span></a></img>
    </tr>
<% end if%>

    <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
	<td height="10"><img src="../Image/space.gif" width="20"></img><img src="../Image/btn.gif"> <a href="ReportBase0010.asp"><span class="pagetitle">舉發案件明細表</span></a></img>
    </tr>
	<tr>
		<td height="44" bgcolor="#FFCC33">3.使用者設定報表</td>
	</tr><%
			For i=0 to UBound(ReportID)
				strSQL="select distinct ReportName from UserRptInfo Where UserId=" & Session("User_ID") & " And UPPER(ReportId)='"&ReportID(i)&"' order by ReportName"
'				response.write strsql
				set rs=conn.execute(strSQL)
				while not rs.eof
					If Not Ifnull(rs("ReportName")) Then
						response.write "<tr bgcolor=""#FFFFFF"" onMouseOver=""this.className='listtitle2'"" onMouseOut=""this.className='listtitle1'""><td height=""10"" ><img src=""../Image/space.gif"" width=""20""></img><img src=""../Image/btn.gif"">"
						response.write "<a href="""&ReportUrl(i)&".asp?ReportName="&rs("ReportName")&"""><span class=""pagetitle"">"&rs("ReportName")&"</span></a>"
						response.write "<input type=""button"" name=""Submit4"" value=""刪除"" onClick=""javascript:if(confirm('是否要刪除?')){funDel('"&Session("User_ID")&"','"&ReportID(i)&"','"&rs("ReportName")&"');}"">"
						response.write "</td></tr>"
					end if
					rs.movenext
				wend
				rs.close
			next
		%>
	</table></td>
  </tr>
  </table>
<input type="hidden" name="dbkind">
<input type="hidden" name="memID">
<input type="hidden" name="reportID">
<input type="hidden" name="reportName">
</form>
</body>
</html>
<script language="javascript">
function funDel(memID,reportID,reportName){
	myForm.dbkind.value="Del";
	myForm.memID.value=memID;
	myForm.reportID.value=reportID;
	myForm.reportName.value=reportName;
	myForm.submit();
}
</script>