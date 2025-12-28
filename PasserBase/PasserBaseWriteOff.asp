<%@ CODEPAGE="65001"%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<%
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

	strSQL="select billsn,(case when cnt=0 then 1 else cnt end) billcnt " &_
		" from (" &_
			"select billsn,count(1) cnt from PasserCreditor pc " &_
				"where Exists(select 'Y' from "&BasSQL&" where sn=pc.billsn) group by billsn" &_
		" ) tmpA"

	filecnt=8
	SN_Cnt=0
	set rs=conn.execute(strSQL)

	While not rs.eof
		
		UitObj.Add rs("billsn") & "_A",cdbl(rs("billcnt"))

		filecnt=filecnt+cdbl(rs("billcnt"))

		If cdbl(rs("billcnt")) > 10 Then
			SN_Cnt=SN_Cnt+10
		else
			SN_Cnt=SN_Cnt+cdbl(rs("billcnt"))
		End if 

		rs.movenext
	Wend
	rs.close

%>
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>hoping</Author>
  <LastAuthor>TitanHsu</LastAuthor>
  <LastPrinted>2022-03-04T07:14:06Z</LastPrinted>
  <Created>2017-06-19T06:52:13Z</Created>
  <LastSaved>2022-03-04T07:06:46Z</LastSaved>
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
  <Style ss:ID="m91775260">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91775280">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0_ "/>
  </Style>
  <Style ss:ID="m91775300">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91775320">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91775340">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91775360">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876864">
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
  <Style ss:ID="m91876884">
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
    ss:Color="#003366" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="m91876640">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876660">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876680">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876700">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876720">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0_);[Red]\(0\)"/>
  </Style>
  <Style ss:ID="m91876740">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876416">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876436">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876456">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876476">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876496">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0_ "/>
  </Style>
  <Style ss:ID="m91876516">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876192">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876212">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876232">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876252">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91876272">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0_);[Red]\(0\)"/>
  </Style>
  <Style ss:ID="m91876292">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="12"
    ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="m91875968">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875988">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91876008">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="m91876028">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="18"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="m91875112">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875132">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875152">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875172">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875192">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="m91875212">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875232">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="m91875252">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
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
  <Style ss:ID="s62">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s63">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s65">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="14"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s66">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s68">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="20"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s104">
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s105">
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
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s106">
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
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <Interior/>
  </Style>
  <Style ss:ID="s107">
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
  <Style ss:ID="s147">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s148">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s149">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s155">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s156">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s157">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s158">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s159">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s160">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"
     ss:Color="#000000"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"
     ss:Color="#000000"/>
   </Borders>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Color="#000000"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s182">
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s189">
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
  <Style ss:ID="s190">
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
  <Style ss:ID="s191">
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
  <Style ss:ID="s192">
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
  <Style ss:ID="s200">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s201">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s202">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="微軟正黑體" x:CharSet="136" x:Family="Swiss" ss:Size="12"
    ss:Color="#000000"/>
  </Style>
  <Style ss:ID="s203">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
  <Style ss:ID="s205">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="16"
    ss:Color="#000000" ss:Bold="1"/>
   <NumberFormat ss:Format="#,##0&quot; &quot;;[Red]&quot;(&quot;#,##0&quot;)&quot;"/>
  </Style>
  <Style ss:ID="s208">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="標楷體" x:CharSet="136" x:Family="Script" ss:Size="15"
    ss:Color="#000000" ss:Bold="1"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="查核清冊(範例)">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='查核清冊(範例)'!R1C1:R28C18"/>
  </Names>
  <Table ss:ExpandedColumnCount="18" ss:ExpandedRowCount="<%=filecnt%>" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s62" ss:DefaultColumnWidth="39.876923076923077"
   ss:DefaultRowHeight="16.061538461538461">
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="79.784615384615378"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="68.123076923076923"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="64.8" ss:Span="1"/>
   <Column ss:Index="5" ss:StyleID="s63" ss:AutoFitWidth="0"
    ss:Width="59.261538461538457"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="66.461538461538467"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="54.830769230769228"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="55.384615384615387"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="57.046153846153842"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="104.12307692307692"/>
   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="70.892307692307696"/>
   <Column ss:StyleID="s62" ss:Width="57.6"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="85.84615384615384"/>
   <Column ss:StyleID="s62" ss:Width="97.476923076923086"/>
   <Column ss:StyleID="s62" ss:Width="88.061538461538476"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="76.984615384615381"/>
   <Column ss:StyleID="s62" ss:Width="57.6"/>
   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="151.20000000000002"/>
   <Row ss:AutoFitHeight="0" ss:Height="22.292307692307695">
    <Cell ss:MergeAcross="1" ss:StyleID="s65"><Data ss:Type="String"><%=year(date)-1911%>年度</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:Index="18" ss:StyleID="s66"><Data ss:Type="String">附表一</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="48.738461538461543">
    <Cell ss:MergeAcross="17" ss:StyleID="s68"><Data ss:Type="String"><%=sys_City%>政府警察局交通罰鍰逾執行時效之債權憑證查核清冊（行政執行事件）</Data><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="50.538461538461533">
    <Cell ss:MergeDown="1" ss:StyleID="m91875112"><Data ss:Type="String">舉發單號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875132"><Data ss:Type="String">單位名稱</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875152"><Data ss:Type="String">違規人&#10;姓名</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875172"><Data ss:Type="String">違規日期</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875192"><Data ss:Type="String">行政罰鍰&#10;金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875212"><Data ss:Type="String">處分書&#10;開立日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875232"><Data ss:Type="String">行政處分&#10;送達日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875252"><Data ss:Type="String">行政處分&#10;確定日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875968"><Data ss:Type="String">初次移送&#10;執行日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91875988"><Data ss:Type="String">債權憑證核發日及文號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m91876008"><Data ss:Type="String">債權憑證&#10;待執行金額</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="6" ss:StyleID="m91876028"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#000000">歷次財產調查及辦理結果</Font><Font
        html:Size="14" html:Color="#000000">(債權憑證再移送日期及文號、歷次財產調查紀錄、結果及債權憑證再移送退案日期及文號等處理情形)</Font></B></ss:Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="32.815384615384616" ss:StyleID="s104">
    <Cell ss:Index="12" ss:StyleID="s105"><Data ss:Type="String">移送日期</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s106"><Data ss:Type="String">移送案號</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s106"><Data ss:Type="String">財產所得查詢日</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s106"><Data ss:Type="String">執行&#10;憑証編號 </Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s106"><Data ss:Type="String">收文文號 </Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s106"><Data ss:Type="String">查詢結果</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s107"><Data ss:Type="String">備註</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <%
   
	strSQL="select SN,BillNo,(select Unitname from unitinfo where unitid=pb.memberstation) UnitName," &_
			"Driver,illegaldate,to_char(illegaldate,'hh24:mi') illegalTime,nvl(forfeit1,0)+nvl(forfeit2,0) forfeit," &_
			"(select judeDate from passerJude where billsn=pb.sn) JudeDate," &_
			"(select max(ArrivedDate) from PasserSendArrived where ArriveType=0) ArrivedDate," &_
			"(select MakeSureDate from PasserSend where billsn=pb.sn) MakeSureDate," &_
			"(select min(SendDate) from PasserSendDetail where billsn=pb.sn) SendDate," &_
			"(select min(PetitionDate) from PasserCreditor where BillSN=pb.sn) PetitionDate," &_
			"(select min(CreditorNumber) from PasserCreditor where BillSN=pb.sn and PetitionDate=(select min(PetitionDate) from PasserCreditor pc2 where BillSN=pb.sn)) CreditorNumber," &_
			"nvl(forfeit1,0)+nvl(forfeit2,0)-(select nvl(sum(PayAmount),0) as PaySum from PasserPay where billsn=pb.sn) noPayAmount " &_
	" from PasserBase pb where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where sn=pb.sn)" &_
	" order by Driver,Billno"

	set rs=conn.execute(strSQL)

	While not rs.eof

		rowCnt=cdbl(trim(UitObj.Item(rs("SN") & "_A")))
		rowCnt=rowCnt-1
		If rowCnt >10 Then rowCnt=9

   %>
   <Row ss:AutoFitHeight="0" ss:StyleID="Default">
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775360"><Data ss:Type="String"><%=trim(rs("BillNo"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775340"><Data ss:Type="String"><%=trim(rs("UnitName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775320"><Data ss:Type="String"><%=trim(rs("Driver"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775300"><Data ss:Type="String"><%=gInitDT(trim(rs("illegaldate")))%>&#10;<%=trim(rs("illegalTime"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775280"><Data ss:Type="Number"><%=trim(rs("forfeit"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91775260"><Data ss:Type="String"><%=gInitDT(trim(rs("judeDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876192"><Data ss:Type="String"><%=gInitDT(trim(rs("ArrivedDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876212"><Data ss:Type="String"><%=gInitDT(trim(rs("MakeSureDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876232"><Data ss:Type="String"><%=gInitDT(trim(rs("SendDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876252"><Data ss:Type="String"><%=gInitDT(trim(rs("PetitionDate")))%>&#10;<%=trim(rs("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876272"><Data ss:Type="Number"><%=trim(rs("noPayAmount"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell><%
	
		strSQL="select (select SendDate from PasserSendDetail where BillSN=pc.Billsn and SN=pc.SendDetailSN) SendDate," &_
				"(select SendNumber from PasserSendDetail where BillSN=pc.Billsn and SN=pc.SendDetailSN) SendNumber," &_
				"PetitionDate,OpenGovNumber,CreditorNumber," &_
				"Decode(CreditorTypeID,1,'無個人財產','清償中') CreditorTypeName " &_
			"from PasserCreditor pc where BillSN="&trim(rs("SN"))&" order by PetitionDate"

		set rspc=conn.execute(strSQL)

		For i = 1 to 10
			If rspc.eof Then exit For 

			If i = 1 Then
	%>
    <Cell ss:StyleID="s147"><Data ss:Type="String"><%=gInitDT(trim(rspc("SendDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s148"><Data ss:Type="String"><%=trim(rspc("SendNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s148"><Data ss:Type="String"><%=gInitDT(trim(rspc("PetitionDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s149"><Data ss:Type="String"><%=trim(rspc("OpenGovNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s149"><Data ss:Type="String"><%=trim(rspc("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s148"><Data ss:Type="String"><%=trim(rspc("CreditorTypeName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeDown="<%=rowCnt%>" ss:StyleID="m91876292"><Data ss:Type="String"></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row><%
			elseif i < cdbl(trim(UitObj.Item(rs("SN") & "_A"))) and i < 10 then
		%>
   <Row ss:AutoFitHeight="0" ss:StyleID="Default">
    <Cell ss:Index="12" ss:StyleID="s155"><Data ss:Type="String"><%=gInitDT(trim(rspc("SendDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s156"><Data ss:Type="String"><%=trim(rspc("SendNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s156"><Data ss:Type="String"><%=gInitDT(trim(rspc("PetitionDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s157"><Data ss:Type="String"><%=trim(rspc("OpenGovNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s157"><Data ss:Type="String"><%=trim(rspc("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s156"><Data ss:Type="String"><%=trim(rspc("CreditorTypeName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row><%		
			else
		%>
   <Row ss:AutoFitHeight="0" ss:Height="16.476923076923075" ss:StyleID="Default">
    <Cell ss:Index="12" ss:StyleID="s158"><Data ss:Type="String"><%=gInitDT(trim(rspc("SendDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s159"><Data ss:Type="String"><%=trim(rspc("SendNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s159"><Data ss:Type="String"><%=gInitDT(trim(rspc("PetitionDate")))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s160"><Data ss:Type="String"><%=trim(rspc("OpenGovNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s160"><Data ss:Type="String"><%=trim(rspc("CreditorNumber"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s159"><Data ss:Type="String"><%=trim(rspc("CreditorTypeName"))%></Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
		<%
			End if 

			rspc.movenext	
		Next
	
		rspc.close

		rs.movenext
	wend
	rs.close
	%>
   <Row ss:AutoFitHeight="0" ss:Height="19.523076923076921">
    <Cell ss:MergeAcross="1" ss:StyleID="m91876864"><Data ss:Type="String">合計</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s190" ss:Formula="=SUM(R[<%="-"&SN_Cnt%>]C:R[-1]C)"><Data
      ss:Type="Number">1200</Data><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s191" ss:Formula="=SUM(R[<%="-"&SN_Cnt%>]C:R[-1]C)"><Data
      ss:Type="Number">1200</Data><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s189"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s192"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="16.476923076923075">
    <Cell ss:MergeAcross="17" ss:StyleID="m91876884"><ss:Data ss:Type="String"
      xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Color="#003366">備註1：本表請覈實填寫並派人自行列管，各單位得自行新增表列欄位，以因應各別內部之需。(※年度以違規日為註記，違規日與裁決日不同年度請另註明。)&#10;</Font><Font
        html:Color="#000000">備註2：各經管人員應善盡善良管理人之注意，定期查調財產且不得延誤移送強制執行。</Font></B></ss:Data><NamedCell
      ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s201"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s201"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s200"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s202"><NamedCell ss:Name="Print_Area"/></Cell>
   </Row>
   <Row ss:AutoFitHeight="0" ss:Height="44.307692307692307">
    <Cell ss:MergeAcross="1" ss:StyleID="s203"><Data ss:Type="String">承辦人：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s205"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><Data ss:Type="String">組長：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="Default"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s205"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="Default"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:MergeAcross="1" ss:StyleID="s208"><Data ss:Type="String">主辦會計：</Data><NamedCell
      ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s203"><NamedCell ss:Name="Print_Area"/></Cell>
    <Cell ss:StyleID="s182"><NamedCell ss:Name="Print_Area"/></Cell>
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
   <Unsynced/>
   <FitToPage/>
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <Scale>58</Scale>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>1</ActiveRow>
     <RangeSelection>R2C1:R2C18</RangeSelection>
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
Response.AddHeader "Content-Disposition", "filename="&Server.UrlPathEncode(fname)&".xls"
response.contenttype="application/x-msexcel; charset=MS950"
%>
