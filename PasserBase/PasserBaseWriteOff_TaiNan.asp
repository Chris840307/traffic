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

	strSQL="select SN,min(to_Number(to_char(illegaldate,'yyyy'))-1911) illegalY" &_
			" from PasserBase pb" &_
				" where Exists(select 'Y' from "&BasSQL&" where sn=pb.sn)" &_
			" group by SN order by illegalY"

	set rs=conn.execute(strSQL)
	filecnt=8:min_year=""
	While not rs.eof
		filecnt=filecnt+1

		If min_year="" Then min_year=rs("illegalY")

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
  <LastPrinted>2022-03-30T07:02:17Z</LastPrinted>
  <Created>2017-06-19T06:52:13Z</Created>
  <LastSaved>2022-03-30T01:06:01Z</LastSaved>
  <Version>14.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>7787</WindowHeight>
  <WindowWidth>20171</WindowWidth>
  <WindowTopX>565</WindowTopX>
  <WindowTopY>576</WindowTopY>
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
  <Style ss:ID="m88944352">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944372">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="m88944392">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="20"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="m88944168">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944188">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944208">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944228">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="m88944248">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944268">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944288">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m88944308">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s65">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s66">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s67">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s78">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s119">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s120">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s121">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s122">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s123">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s125">
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s127">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="14"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s129">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="20"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s146">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s149">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="15"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s153">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s181">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="14"
    ss:Color="#000000"/>
   <Interior/>
  </Style>
  <Style ss:ID="s182">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s183">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s184">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s201">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s202">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s203">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="11"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s204">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s205">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s206">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="9"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s207">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#FF0000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s208">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s209">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s210">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s211">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s212">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s213">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s214">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s215">
   <Alignment ss:Horizontal="Left" ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s216">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s217">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s218">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s219">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s220">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s221">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s222">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:FontName="新細明體" x:CharSet="136" x:Family="Roman" ss:Size="8"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s226">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s227">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s228">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
 </Styles>
 <Worksheet ss:Name="查核清冊(範例)">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='查核清冊(範例)'!R1C1:R10C33"/>
  </Names>
  <Table ss:ExpandedColumnCount="34" ss:ExpandedRowCount="<%=filecnt%>" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s65" ss:DefaultColumnWidth="39.876923076923077"
   ss:DefaultRowHeight="16.061538461538461">
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="27.138461538461538"/>
   <Column ss:StyleID="s65" ss:Width="56.492307692307691"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="63.692307692307693"/>
   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="42.092307692307692"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="49.846153846153847"/>
   <Column ss:StyleID="s65" ss:Width="47.07692307692308"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="44.307692307692307"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="50.953846153846158"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="62.030769230769238"/>
   <Column ss:StyleID="s66" ss:AutoFitWidth="0" ss:Width="44.307692307692307"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="38.215384615384615"
    ss:Span="21"/>
   <Column ss:Index="33" ss:StyleID="s65" ss:AutoFitWidth="0"
    ss:Width="61.476923076923072"/>
   <Row ss:Height="22.153846153846153" ss:StyleID="Default">
    <Cell ss:MergeAcross="1" ss:StyleID="s127"><Data ss:Type="String"><%=trim(min_year)%>年度</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s66"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s66"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s67"><Data ss:Type="String">附表一</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="48.738461538461543" ss:StyleID="Default">
    <Cell ss:MergeAcross="32" ss:StyleID="s129"><Data ss:Type="String"><%=replace(trim(sys_City),"台","臺")%>政府警察局交通罰鍰逾執行時效之債權憑證查核清冊（行政執行事件）</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="34.061538461538461" ss:StyleID="Default">
    <Cell ss:MergeDown="1" ss:StyleID="m88944168"><Data ss:Type="String">案件&#10;編號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944188"><Data ss:Type="String">單位名稱</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944208"><Data ss:Type="String">案名&#10;（處分書編號）及違規日期</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944228"><Data ss:Type="String">行政罰鍰&#10;金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944248"><Data ss:Type="String">處分書&#10;開立日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944268"><Data ss:Type="String">行政處分&#10;送達日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944288"><Data ss:Type="String">行政處分&#10;確定日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944308"><Data ss:Type="String">初次移送&#10;執行日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944352"><Data ss:Type="String">債權憑證核發日及文號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m88944372"><Data ss:Type="String">債權憑證&#10;待執行金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="22" ss:StyleID="m88944392"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Size="18"
        html:Color="#000000">歷次財產調查及辦理結果</Font><Font html:Size="14"
        html:Color="#000000">(債權憑證再移送日期及文號、歷次財產調查紀錄、結果及債權憑證再移送退案日期及文號等處理情形)</Font></B></ss:Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"/>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="33.646153846153844" ss:StyleID="s78">
	<%
		For i = 0 to 10
			If i = 0 Then
			%>
    <Cell ss:Index="11" ss:MergeAcross="1" ss:StyleID="s226"><Data ss:Type="String"><%=cdbl(min_year)+i%>年</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
			<%
			else
		%>
    <Cell ss:MergeAcross="1" ss:StyleID="s227"><Data ss:Type="String"><%=cdbl(min_year)+i%>年</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
		<%		
			End if
		Next	
	%>
    <Cell ss:StyleID="s228"><Data ss:Type="String">是否善盡管理責任</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <%
   
	strSQL="select SN,BillNo,(select Unitname from unitinfo where unitid=pb.memberstation) UnitName," &_
			"Driver,illegaldate,nvl(forfeit1,0)+nvl(forfeit2,0) forfeit," &_
			"(select judeDate from passerJude where billsn=pb.sn) JudeDate," &_
			"(select max(ArrivedDate) from PasserSendArrived where ArriveType=0 and passersn=pb.sn) ArrivedDate," &_
			"(select max(Note) from PasserSendArrived where ArriveType=0 and passersn=pb.sn) ArriveNote," &_
			"(select MakeSureDate from PasserSend where billsn=pb.sn) MakeSureDate," &_
			"(select min(SendDate) from PasserSendDetail where billsn=pb.sn) SendDate," &_
			"(select min(OpenGovNumber) from PasserSendDetail where billsn=pb.sn and SendDate=" &_
				"(select min(SendDate) from PasserSendDetail psd where billsn=pb.sn)" &_
			") SendOpenGovNumber," &_
			"(select min(PetitionDate) from PasserCreditor where BillSN=pb.sn) PetitionDate," &_
			"(select min(CreditorNumber) from PasserCreditor where BillSN=pb.sn and PetitionDate=" &_
				"(select min(PetitionDate) from PasserCreditor pc2 where BillSN=pb.sn)" &_
			") CreditorNumber," &_
			"nvl(forfeit1,0)+nvl(forfeit2,0)-(select nvl(sum(PayAmount),0) as PaySum from PasserPay where billsn=pb.sn) noPayAmount " &_
	" from PasserBase pb where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where sn=pb.sn)" &_
	" order by Driver,Billno"

	set rs=conn.execute(strSQL)
	syscnt=0
While not rs.eof
	syscnt=syscnt+1
   %>
   <Row ss:AutoFitHeight="0" ss:Height="90.553846153846152" ss:StyleID="Default">
    <Cell ss:StyleID="s201"><Data ss:Type="Number"><%=trim(syscnt)%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s202"><Data ss:Type="String"><%=trim(rs("UnitName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><Data ss:Type="String"><%=trim(rs("Driver"))%>&#10;<%=trim(rs("BillNo"))%>&#10;<%=gInitDT(trim(rs("illegaldate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s204"><Data ss:Type="Number"><%=trim(rs("forfeit"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s205"><Data ss:Type="Number"><%=gInitDT(trim(rs("JudeDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s206"><Data ss:Type="String"><%=gInitDT(trim(rs("ArrivedDate")))%>&#10;-<%=rs("ArriveNote")%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s207"><Data ss:Type="String"><%=gInitDT(trim(rs("MakeSureDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s208"><Data ss:Type="String"><%=gInitDT(trim(rs("SendDate")))%>&#10;<%=trim(rs("SendOpenGovNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><Data ss:Type="String"><%=gInitDT(trim(rs("PetitionDate")))%>&#10;<%=trim(rs("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s204"><Data ss:Type="Number"><%=trim(rs("noPayAmount"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
	<%
	For i = 0 to 10
		strSQL="select min(senddate) senddate,min(OpenGovNumber) SendOpenGovNumber from PasserSendDetail where billsn="&trim(rs("SN"))&" and to_char(senddate,'yyyy')='"&cdbl(min_year)+i+1911&"'"

		set rsc=conn.execute(strSQL)
		If not rsc.eof Then
		%>
    <Cell ss:StyleID="s217"><Data ss:Type="String"><%=gInitDT(trim(rsc("senddate")))%>&#10;<%=trim(rsc("SendOpenGovNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>	
		<%

		else
		%>
    <Cell ss:StyleID="s217"><Data ss:Type="String">&#10;</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
		<%
		End if 
		rsc.close
	'=============================================================
		strSQL="select min(PetitionDate) PetitionDate,min(CreditorNumber) CreditorNumber from PasserCreditor where billsn="&trim(rs("SN"))&" and to_char(PetitionDate,'yyyy')='"&cdbl(min_year)+i+1911&"'"

		set rsc=conn.execute(strSQL)
		If not rsc.eof Then
		%>
    <Cell ss:StyleID="s218"><Data ss:Type="String"><%=gInitDT(trim(rsc("PetitionDate")))%>&#10;<%=trim(rsc("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>		
		<%
		else
		%>
    <Cell ss:StyleID="s218"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
		<%
		End if 
		rsc.close
	next
	%>
    <Cell ss:StyleID="s218"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s65"/>
   </Row>	
	<%

	rs.movenext
wend
rs.close
	%>
   <Row ss:AutoFitHeight="0" ss:Height="48.738461538461543" ss:StyleID="Default">
    <Cell ss:MergeAcross="1" ss:StyleID="s181"><Data ss:Type="String">合計</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s183" ss:Formula="=SUM(R[-<%=filecnt-8%>]C:R[-1]C)"><Data ss:Type="Number"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s153" ss:Formula="=SUM(R[-<%=filecnt-8%>]C:R[-1]C)"><Data ss:Type="Number"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s184"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="39.046153846153842" ss:StyleID="Default">
    <Cell ss:MergeAcross="32" ss:StyleID="s146"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#16365C">備註1：本表請覈實填寫並派人自行列管，各單位得自行新增表列欄位，以因應各別內部之需。(※年度以違規日為註記，違規日與裁決日不同年度請另註明。)&#10;</Font><Font
        html:Color="#000000">備註2：各經管人員應善盡善良管理人之注意，定期查調財產且不得延誤移送強制執行。</Font></B></ss:Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="39.046153846153842" ss:StyleID="Default">
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s120"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s120"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s119"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s121"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="39.046153846153842" ss:StyleID="s125">
    <Cell ss:MergeAcross="1" ss:StyleID="s122"><Data ss:Type="String">承辦人：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s123"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><Data ss:Type="String">組長：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="Default"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s123"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="Default"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s149"><Data ss:Type="String">主辦會計：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="Default"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="3" ss:StyleID="s122"><Data ss:Type="String">分局長：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s122"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Layout x:Orientation="Landscape" x:CenterHorizontal="1"/>
    <Header x:Margin="0.31496062992126012"/>
    <Footer x:Margin="0.31496062992126012" x:Data="&amp;C第 &amp;P 頁，共 &amp;N 頁"/>
    <PageMargins x:Bottom="0.74803149606299213" x:Left="0.39370078740157505"
     x:Right="0.19685039370078702" x:Top="0.74803149606299213"/>
   </PageSetup>
   <FitToPage/>
   <Print>
    <FitHeight>0</FitHeight>
    <ValidPrinterInfo/>
    <PaperSizeIndex>8</PaperSizeIndex>
    <Scale>81</Scale>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>1</ActiveRow>
     <RangeSelection>R2C1:R2C33</RangeSelection>
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
fname=year(now)&fMnoth&fDay&"_慢車註銷案件清冊"
Response.AddHeader "Content-Disposition", "filename="&Server.Urlencode(fname)&".xls"
response.contenttype="application/x-msexcel; charset=MS950"
%>
