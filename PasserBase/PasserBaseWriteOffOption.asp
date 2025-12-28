<%@ CODEPAGE="65001"%>
<%Response.CodePage=65001
Response.Charset="UTF-8"%><!-- #include file="../Common/ReportDbUtil.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"--><%
	Server.ScriptTimeout=6000
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	sys_City=replace(sys_City,"台中縣","台中市")
	sys_City=replace(sys_City,"台南縣","台南市")

	showCreditor=false
	if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
		showCreditor=true
	end If 

	If Not ifnull(request("Sys_SendBillSN")) Then

		sys_billsn=request("Sys_SendBillSN")
	elseif Not ifnull(request("hd_BillSN")) Then

		sys_billsn=request("hd_BillSN")
	else

		sys_billsn=request("BillSN")
	End If 

	tmp_billsn=split(sys_billsn,",")

	sys_billsn=""

	For i = 0 to Ubound(tmp_billsn)

		If i >0 then

			If i mod 100 = 0 Then

				sys_billsn=sys_billsn&"@"
			elseif sys_billsn<>"" then

				sys_billsn=sys_billsn&","
			end If 
		end if

		sys_billsn=sys_billsn&tmp_billsn(i)

	Next

	tmpSQL=""

	If Ubound(tmp_billsn) >= 100 Then

		sys_billsn=split(sys_billsn,"@")
		
		For i = 0 to Ubound(sys_billsn)
			
			If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
			
			tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
		Next

	else

		tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

	End if 

	BasSQL="("&tmpSQL&") tmpPasser"

	
	Set UitObj = Server.CreateObject("Scripting.Dictionary")

	strSQL="select count(1) cnt from PasserBase pb " &_
			"where Exists(select 'Y' from "&BasSQL&" where sn=pb.sn) "

	SN_Cnt=0
	
	set rs=conn.execute(strSQL)

	While not rs.eof
		
		SN_Cnt=cdbl(rs("cnt"))

		rs.movenext
	Wend
	rs.close

%><?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>hoping</Author>
  <LastAuthor>TitanHsu</LastAuthor>
  <LastPrinted>2022-03-09T02:39:49Z</LastPrinted>
  <Created>2018-03-08T01:36:28Z</Created>
  <LastSaved>2022-03-09T02:49:10Z</LastSaved>
  <Version>14.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>6569</WindowHeight>
  <WindowWidth>17247</WindowWidth>
  <WindowTopX>388</WindowTopX>
  <WindowTopY>554</WindowTopY>
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
  <Style ss:ID="s65" ss:Name="一般 4">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
  </Style>
  <Style ss:ID="s69" ss:Name="千分位 2">
   <NumberFormat ss:Format="&quot; &quot;#,##0.00&quot; &quot;;&quot;-&quot;#,##0.00&quot; &quot;;&quot; -&quot;00&quot; &quot;;&quot; &quot;@&quot; &quot;"/>
  </Style>
  <Style ss:ID="s72" ss:Name="千分位[0] 4">
   <NumberFormat ss:Format="&quot; &quot;#,##0&quot; &quot;;&quot;-&quot;#,##0&quot; &quot;;&quot; - &quot;;&quot; &quot;@&quot; &quot;"/>
  </Style>
  <Style ss:ID="s114" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s115" ss:Parent="s69">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s117" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s118" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="0&quot; &quot;;[Red]&quot;(&quot;0&quot;)&quot;"/>
   <Protection/>
  </Style>
  <Style ss:ID="s119" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="@"/>
   <Protection/>
  </Style>
  <Style ss:ID="s120" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s121" ss:Parent="s72">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s122" ss:Parent="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s130" ss:Parent="s65">
   <Alignment ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="11"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s131" ss:Parent="s65">
   <Alignment ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s132" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s134" ss:Parent="s65">
   <Alignment ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="0&quot; &quot;;[Red]&quot;(&quot;0&quot;)&quot;"/>
   <Protection/>
  </Style>
  <Style ss:ID="s137" ss:Parent="s65">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="11"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s138" ss:Parent="s65">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s139" ss:Parent="s65">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
   <Protection/>
  </Style>
  <Style ss:ID="s140" ss:Parent="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s141" ss:Parent="s65">
   <Alignment ss:Horizontal="Right" ss:Vertical="Top"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s142" ss:Parent="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
   <Protection/>
  </Style>
  <Style ss:ID="s143">
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s146" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000"/>
   <Interior/>
   <Protection/>
  </Style>
  <Style ss:ID="s150" ss:Parent="s69">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s151" ss:Parent="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior/>
   <NumberFormat ss:Format="0&quot; &quot;;[Red]&quot;(&quot;0&quot;)&quot;"/>
   <Protection/>
  </Style>
  <Style ss:ID="s154">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior/>
  </Style>
 </Styles>
 <Worksheet ss:Name="範例">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="=範例!R1C1:R30C14"/>
  </Names>
  <Table ss:ExpandedColumnCount="14" ss:ExpandedRowCount="<%=cdbl(SN_Cnt)+10%>" x:FullColumns="1"
   x:FullRows="1" ss:DefaultColumnWidth="39.876923076923077"
   ss:DefaultRowHeight="16.476923076923075">
   <Column ss:AutoFitWidth="0" ss:Width="14.4"/>
   <Column ss:AutoFitWidth="0" ss:Width="47.630769230769232"/>
   <Column ss:Width="22.153846153846153"/>
   <Column ss:AutoFitWidth="0" ss:Width="34.338461538461537"/>
   <Column ss:Width="37.661538461538463"/>
   <Column ss:AutoFitWidth="0" ss:Width="40.430769230769229"/>
   <Column ss:AutoFitWidth="0" ss:Width="50.953846153846158"/>
   <Column ss:StyleID="s143" ss:AutoFitWidth="0" ss:Width="29.907692307692308"/>
   <Column ss:StyleID="s143" ss:AutoFitWidth="0" ss:Width="31.569230769230771"/>
   <Column ss:AutoFitWidth="0" ss:Width="23.815384615384616"/>
   <Column ss:AutoFitWidth="0" ss:Width="15.507692307692309"/>
   <Column ss:AutoFitWidth="0" ss:Width="38.215384615384615"/>
   <Column ss:AutoFitWidth="0" ss:Width="19.938461538461539"/>
   <Column ss:AutoFitWidth="0" ss:Width="74.215384615384608"/>
   <Row ss:Height="22.153846153846153">
    <Cell ss:MergeAcross="13" ss:StyleID="s146"><Data ss:Type="String"><%=replace(trim(sys_City),"台","臺")%>政府警察局註銷應收款項清冊</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">序號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">單位&#10;名稱</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">年度</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">確定日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">舉發單號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">違規人&#10;姓名</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s114"><Data ss:Type="String">身分證號碼</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s150"><Data ss:Type="String">保留&#10;金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s150"><Data ss:Type="String">註銷&#10;金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="s114"><Data ss:Type="String">註銷原因</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="s151"><Data ss:Type="String">移送(發文)案號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="103.84615384615384">
    <Cell ss:Index="10" ss:StyleID="s114"><Data ss:Type="String">已取得首次債權憑證</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String">帳載錯誤</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String">應收款列帳已逾四年尚無法收繳，且尚未取得債權憑證</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String">其他特殊情形</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row><%
		strSQL="select SN,(select Unitname from unitinfo where unitid=pb.memberstation) UnitName," &_
			"to_char(illegaldate,'yyyy')-1911 illegalYear," &_
			"to_char(illegaldate,'yyyy')-1911||to_char(illegaldate,'mmdd') illegalDate," &_
			"(select to_char(MakeSureDate,'yyyy')-1911||to_char(MakeSureDate,'mmdd') from passerSend where billsn=pb.sn) MakeSureDate," &_
			"BillNo,Driver,DriverID," &_
			"nvl(forfeit1,0)+nvl(forfeit2,0) forfeit," &_
			"nvl(forfeit1,0)+nvl(forfeit2,0)-(select nvl(sum(PayAmount),0) as PaySum from PasserPay where billsn=pb.sn) noPayAmount," &_
			"(select min(PetitionDate) from PasserCreditor where BillSN=pb.sn) PetitionDate," &_
			"(select min(CreditorNumber) from PasserCreditor where BillSN=pb.sn and PetitionDate=" &_
				"(select min(PetitionDate) from PasserCreditor pc2 where BillSN=pb.sn)" &_
			") CreditorNumber " &_
		" from PasserBase pb where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where sn=pb.sn)" &_
		" order by illegalYear,UnitName,Billno"
		
	cntfile=0
	t_sum01=0:t_sum02=0
	set rs=conn.execute(strSQL)
	While not rs.eof
		

		cntfile=cntfile+1

		
		t_sum01=t_sum01+cdbl(rs("forfeit"))
		t_sum02=t_sum02+cdbl(rs("noPayAmount"))
   %>
   <Row ss:AutoFitHeight="0" ss:Height="22.153846153846153">
    <Cell ss:StyleID="s117"><Data ss:Type="String"><%=trim(cntfile)%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String"><%=trim(rs("UnitName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String"><%=trim(rs("illegalYear"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s118"><Data ss:Type="String"><%=trim(rs("MakeSureDate"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><Data ss:Type="String"><%=trim(rs("BillNo"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s120"><Data ss:Type="String"><%=trim(rs("Driver"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><Data ss:Type="String"><%=trim(rs("DriverID"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s121"><Data ss:Type="Number"><%=cdbl(rs("forfeit"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s121"><Data ss:Type="Number"><%=cdbl(rs("noPayAmount"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String">ˇ</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><Data ss:Type="String"><%=gInitDT(trim(rs("PetitionDate")))%>&#10;<%=trim(rs("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <%
   
		rs.movenext
	Wend
	rs.close
   %>
   <Row ss:AutoFitHeight="0" ss:Height="30.599999999999998">
    <Cell ss:StyleID="s130"><Data ss:Type="String">總計</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s131"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s131"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s132"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s120"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s120"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s115"><Data ss:Type="Number"><%=cdbl(t_sum01)%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s115"><Data ss:Type="Number"><%=cdbl(t_sum02)%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s114"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s134"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="31.569230769230771">
    <Cell ss:MergeAcross="13" ss:StyleID="s154"><Data ss:Type="String">承辦人：              單位主管：               主辦會計:                機關長官:</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:StyleID="s137"><Data ss:Type="String">填表注意事項：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s139"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s139"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s138"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:StyleID="s141"><Data ss:Type="String">1.</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><Data ss:Type="String">本表「年度」欄按年度順序列示，每一年度結一小計，各年度小計結一合計。</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:StyleID="s141"><Data ss:Type="String">2.</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><Data ss:Type="String">註銷原因之選定，請以「ˇ」符號勾選。</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:StyleID="s141"><Data ss:Type="String">3.</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><Data ss:Type="String">『保留金額』攔位如案件已有部分金額收繳時，應以扣除收繳金額數填列，『註銷金額』攔位則依前攔保留金額攔位金額填列。</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s142"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s140"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:Height="16.061538461538461">
    <Cell ss:StyleID="s141"><Data ss:Type="String">4.</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="12" ss:StyleID="s140"><Data ss:Type="String">所檢附證明文件影本請依序排列，俾利核對。</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Layout x:CenterHorizontal="1"/>
    <Header x:Margin="0.11811023622047202"/>
    <Footer x:Margin="0.11811023622047202" x:Data="&amp;C第 &amp;P 頁，共 &amp;N 頁"/>
    <PageMargins x:Bottom="0.55118110236220408" x:Left="0.11811023622047202"
     x:Right="0.11811023622047202" x:Top="0.55118110236220408"/>
   </PageSetup>
   <Print>
    <FitWidth>0</FitWidth>
    <FitHeight>0</FitHeight>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>1</ActiveRow>
     <ActiveCol>5</ActiveCol>
     <RangeSelection>R2C6:R3C6</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
<%
ConnExecute Request.ServerVariables ("SCRIPT_NAME")&"||"& DateDiff("s",log_start,now) &"||"& startDate_q & "~" & endDate_q ,361

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_應收款項註銷清冊"
Response.AddHeader "Content-Disposition", "filename="&Server.Urlencode(fname)&".xls"
response.contenttype="application/x-msexcel; charset=MS950"
%>
