<%@ CODEPAGE="65001"%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<%
Server.ScriptTimeout=6000
log_start=now
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=31"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then 
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenCity=replace(trim(rsUInfo("value")),"台","臺")
	end if
end if 
rsUInfo.close
set rsUInfo=nothing

sql = "select Value from Apconfigure where ID=35"
Set RSSystem = Conn.Execute(sql)
if Not RSSystem.Eof Then
	rptHead1 = RSSystem("Value")
End If 

RSSystem.close

strUit=split(",JM00,JS00,JO00,JQ00,JN00,JP00,JR00,JT00",",")

ArgueDate1=year(now)&"/"&(month(now)-1)&"/01 0:0:0"
ArgueDate2=year(now)&"/"&(month(now)-1)&"/"&day(dateAdd("d",-1,year(now)&"/"&(month(now))&"/01"))&" 23:59:59"

If request("PayDate1") <> "" Then
	
	ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
	ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"
End if 



strwhere=" and PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and payno is not null"

whereUnit=""

if request("Sys_MemberStation")<>"" then

	strwhere=strwhere&" and exists(select 'Y' from passerbase where MemberStation in('"&request("Sys_MemberStation")&"') and recordstateid=0 and sn=passerpay.billsn)"

	whereUnit=" and unitid in('"&request("Sys_MemberStation")&"')"

else

	strwhere=strwhere&" and exists(select 'Y' from passerbase where MemberStation in(select UnitID from UnitInfo where UnitLevelID=2 and UnitName like '%分局') and recordstateid=0 and sn=passerpay.billsn)"
end If 

titleUnit="(總表)"

Set arrCnt = Server.CreateObject("Scripting.Dictionary")

if request("Sys_MemberStation")<>"" then

	SqlUit = "select UnitName from UnitInfo where UnitLevelID=2 and UnitName like '%分局'"&whereUnit
	set rsuit=conn.execute(SqlUit)
	If not rsuit.eof Then
		titleUnit=rsuit("UnitName")
	End if 

	rsuit.close
end If 

strSQL="select to_char(paydate,'YYYY') years,to_char(paydate,'MM') months,to_char(paydate,'DD') days," & _
"min(payno) minpayno,max(payno) maxpayno,1 payType,count(1) cnt " & _
" from passerpay where payno is not null" & strwhere & _
" group by to_char(paydate,'YYYY'),to_char(paydate,'MM'),to_char(paydate,'DD')"

strSQL=strSQL&" union all "

strSQL=strSQL&"select to_char(paydate,'YYYY') years,to_char(paydate,'MM') months,to_char(paydate,'DD') days," & _
"payno minpayno,payno maxpayno,2 payType,1 cnt " & _
" from PASSERPAYDEL where payno is not null" & replace(replace(strwhere,"passerpay","PASSERPAYDEL"),"and recordstateid=0","") & _
" order by years,months,days"

set rs=conn.execute(strSQL)
tmpdate="":chkdate="":tableCnt=0
While not rs.eof

	tmpdate=rs("years")&"_"&rs("months")&"_"&rs("days")

	If instr(chkdate,tmpdate) <=0 Then

		If chkdate <> "" Then chkdate=chkdate&","
		chkdate=chkdate & tmpdate

		arrCnt.Add tmpdate &"_Y","" & rs("years") & ""
		arrCnt.Add tmpdate &"_M","" & rs("months") & ""
		arrCnt.Add tmpdate &"_D","" & rs("days") & ""
		arrCnt.Add tmpdate &"_1_min",""""
		arrCnt.Add tmpdate &"_1_max",""""
		arrCnt.Add tmpdate &"_1_cnt",0
		arrCnt.Add tmpdate &"_2_min"," "
		arrCnt.Add tmpdate &"_2_max"," "
		arrCnt.Add tmpdate &"_2_cnt",0
	End if 
	If rs("payType") = 1 Then

		arrCnt.Item(tmpdate& "_" &rs("payType")&"_min")=rs("minpayno")
		arrCnt.Item(tmpdate& "_" &rs("payType")&"_max")=rs("maxpayno")
		arrCnt.Item(tmpdate& "_" &rs("payType")&"_cnt")=rs("cnt")
	else
		If arrCnt.Item(tmpdate& "_" &rs("payType")&"_min") <>"" Then 

			arrCnt.Item(tmpdate& "_" &rs("payType")&"_min")=arrCnt.Item(tmpdate& "_" &rs("payType")&"_min")&","
		End if 
		
		arrCnt.Item(tmpdate& "_" &rs("payType")&"_min")=arrCnt.Item(tmpdate& "_" &rs("payType")&"_min")&rs("minpayno")
		arrCnt.Item(tmpdate& "_" &rs("payType")&"_cnt")=arrCnt.Item(tmpdate& "_" &rs("payType")&"_cnt")+cdbl(rs("cnt"))
	End if 

	rs.movenext
Wend

rs.close

arrdate=split(chkdate,",")
tableCnt=Ubound(arrdate)+1

If chkdate = "" Then
	Response.End
End if 
%>
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>TitanHsu</Author>
  <LastAuthor>TitanHsu</LastAuthor>
  <LastPrinted>2020-08-03T07:44:14Z</LastPrinted>
  <Created>2020-08-03T07:43:26Z</Created>
  <Version>14.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>6425</WindowHeight>
  <WindowWidth>11642</WindowWidth>
  <WindowTopX>443</WindowTopX>
  <WindowTopY>44</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62" ss:Name="一般 2">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s65" ss:Parent="s62">
   <Alignment ss:Horizontal="CenterAcrossSelection" ss:Vertical="Center"/>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s68" ss:Parent="s62">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s69" ss:Parent="s62">
   <Alignment ss:Horizontal="Distributed" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s70" ss:Parent="s62">
   <Alignment ss:Horizontal="Distributed" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="11"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s71" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="11"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s73" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="11"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="General&quot;份&quot;"/>
  </Style>
  <Style ss:ID="s77" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s78" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Times New Roman" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s124" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s136" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s139" ss:Parent="s62">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="工作表1">
  <Table ss:ExpandedColumnCount="12" ss:ExpandedRowCount="<%=tableCnt+8%>" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="48.738461538461543"
   ss:DefaultRowHeight="16.061538461538461">
   <Row ss:Height="22.153846153846153">
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:MergeAcross="5" ss:StyleID="s124"><Data ss:Type="String"><%=rptHead1&titleUnit%></Data></Cell>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
   </Row>
   <Row ss:Height="22.153846153846153">
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:MergeAcross="5" ss:StyleID="s124"><Data ss:Type="String">自行收納款項收據紀錄卡</Data></Cell>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
    <Cell ss:StyleID="s65"/>
   </Row>
   <Row ss:Height="16.615384615384617">
    <Cell ss:MergeAcross="2" ss:StyleID="s139"><Data ss:Type="String">費別：交通罰鍰收入</Data></Cell>
    <Cell ss:MergeAcross="5" ss:StyleID="s136"><Data ss:Type="String"><%
     Response.Write "中華民國"&(cdbl(arrCnt.Item(arrdate(0)& "_Y"))-1911)
     Response.Write "年度"&arrCnt.Item(arrdate(0)& "_M")&"月份"
    %></Data></Cell>
    <Cell ss:StyleID="s68"/>
    <Cell ss:StyleID="s68"/>
    <Cell ss:StyleID="s68"/>
   </Row>
   <Row ss:Height="16.615384615384617">
    <Cell ss:MergeAcross="8" ss:StyleID="s69"><Data ss:Type="String">銷號數</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="s69"><Data ss:Type="String">結存數</Data></Cell>
   </Row>
   <Row ss:Height="16.615384615384617">
    <Cell ss:MergeAcross="2" ss:StyleID="s69"><Data ss:Type="String">銷號日期</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s69"><Data ss:Type="String">起訖號數</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s69"><Data ss:Type="String">份數</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s69"><Data ss:Type="String">作廢收據</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s69"><Data ss:Type="String">銷號人核章</Data></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s69"><Data ss:Type="String">起訖號數</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s69"><Data ss:Type="String">份數</Data></Cell>
   </Row>
   <Row ss:Height="16.615384615384617">
    <Cell ss:StyleID="s69"><Data ss:Type="String">年</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">月</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">日</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">起</Data></Cell>
    <Cell ss:StyleID="s69"><Data ss:Type="String">訖</Data></Cell>
    <Cell ss:Index="7" ss:StyleID="s70"><Data ss:Type="String">號數</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">份數</Data></Cell>
    <Cell ss:Index="10" ss:StyleID="s69"><Data ss:Type="String">起</Data></Cell>
    <Cell ss:StyleID="s77"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><Font html:Color="#000000"> </Font><Font
       html:Face="標楷體" x:CharSet="136" x:Family="Script" html:Color="#000000">訖</Font></ss:Data></Cell>
   </Row>
   <%
   For i = 0 to Ubound(arrdate)
	Response.Write "<Row ss:Height=""16.615384615384617"">"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""><Data ss:Type=""Number"">"&cdbl(arrCnt.Item(arrdate(i)& "_Y"))-1911&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""><Data ss:Type=""Number"">"&arrCnt.Item(arrdate(i)& "_M")&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""><Data ss:Type=""Number"">"&arrCnt.Item(arrdate(i)& "_D")&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s78""><Data ss:Type=""String"">"&arrCnt.Item(arrdate(i)& "_1_min")&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s78""><Data ss:Type=""String"">"&arrCnt.Item(arrdate(i)& "_1_max")&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s73""><Data ss:Type=""Number"">"
	If arrCnt.Item(arrdate(i)& "_1_cnt") > 0 Then
		Response.Write arrCnt.Item(arrdate(i)& "_1_cnt")
	End if 
	Response.Write "</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s78""><Data ss:Type=""String"">"&arrCnt.Item(arrdate(i)& "_2_min")&"</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s73""><Data ss:Type=""Number"">"
	If arrCnt.Item(arrdate(i)& "_2_cnt") > 0 Then
		Response.Write arrCnt.Item(arrdate(i)& "_2_cnt")
	End if 
	Response.Write "</Data></Cell>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""/>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""/>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s71""/>"&vbcrlf
	Response.Write "<Cell ss:StyleID=""s73""/>"&vbcrlf
	Response.Write "</Row>"&vbcrlf
   Next
   %>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>10</ActiveRow>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="工作表2">
  <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="48.738461538461543"
   ss:DefaultRowHeight="16.061538461538461">
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="工作表3">
  <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="48.738461538461543"
   ss:DefaultRowHeight="16.061538461538461">
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_收據紀錄卡"
Response.AddHeader "Content-Disposition", "filename="&Server.UrlPathEncode(fname)&".xls"
response.contenttype="application/x-msexcel; charset=utf-8"

conn.close
set conn=nothing
%>