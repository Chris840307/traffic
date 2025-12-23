	<!--#include virtual="traffic/Common/DB.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<%
	tmpSql="Select value from apconfigure where id=31"
  	set rscnt=conn.execute(tmpSql)
	 Sys_City=rscnt("value") 
    set rscnt=nothing



Server.ScriptTimeout = 8648000
unit = Request("unit")
UnitID_q = Request("UnitID_q")
UnitID_Title = Request("UnitID_Title")
 'DateType="ILlEGALDate"
 DateType=Request("DateType")
if DateType="ILLEGALDate" then 
  DateName="(違規日)"
elseif  DateType="BillFillDate" then 
  DateName="(填單日)"
elseif  DateType="RecordDate" then 
  DateName="(建檔日)"
end if 
             date1=gOutDT(Trim(Request("startDate_q")))
             date2=gOutDT(Trim(Request("endDate_q")))
tmpSql=" and RecordStateID='0' "

unitList=trim(request("UnitID_q"))

unitList = Split(unitList,"~")
filename=""
If unit="y" Then
	Sys_UnitID=""
	for i=0 to UBound(unitList)
		if Sys_UnitID<>"" then Sys_UnitID=Sys_UnitID&"','"
		Sys_UnitID=Sys_UnitID&unitList(i)


		If ifnull(filename) Then 
			strSQL="select unitname from unitinfo where unitid='"&unitList(i)&"'"
			set rsuit=conn.execute(strSQL)
			filename=rsuit("unitname")
			rsuit.close
		end if

	next
	tmpSql = tmpSql & "  and Billno is not null and BillUnitId in (select unitid from unitinfo where UnitTypeID in ('" & Sys_UnitID & "') "
	tmpSql = tmpSql & " union select unitid from unitinfo where unitid in ('" & Sys_UnitID & "')) "

elseif unit="n" then 

  tmpSql = tmpSql & " and Billno is not null "
  
End If


strSQL="select BillTypeID,rule1,Total from (Select BillTypeID,rule1,count(billno) as Total from BillBase "
strSQL=strSQL&" where  CarSimpleID in ('2','1')  and (CarAddID not in ('9','5') or CarAddID is null)"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule1"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule2 as rule1,count(billno) as Total from BillBase  where CarSimpleID in ('2','1')"
strSQL=strSQL&" and (CarAddID not in ('9','5') or CarAddID is null)"
strSQL=strSQL&" and rule2 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule2"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule3 as rule1,count(billno) as Total from BillBase  where   CarSimpleID in ('2','1')"
strSQL=strSQL&" and (CarAddID not in ('9','5') or CarAddID is null)"
strSQL=strSQL&" and rule3 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&"  group by BillTypeID,rule3"
strSQL=strSQL&"  )"
strSQL=strSQL&"  order by 2"
'response.write strsql
'response.end
set rs2=conn.execute(strSQL)
'汽車
	E18=0 : F18=0 :	G18=0 :	H18=0 :	I18=0 :	J18=0 :	K18=0 :	L18=0 :	M18=0 :	N18=0 :	O18=0 :	P18=0 :	Q18=0 :	R18=0 :	S18=0 :	T18=0 :	U18=0 :	V18=0
	W18=0 :	X18=0 :	Y18=0 :	Z18=0 : AA18=0 :	AB18=0 :	AC18=0 : 	AD18=0 :	AE18=0
	E19=0 : F19=0 :	G19=0 :	H19=0 :	I19=0 :	J19=0 :	K19=0 :	L19=0 :	M19=0 :	N19=0 :	O19=0 :	P19=0 :	Q19=0 :	R19=0 :	S19=0 :	T19=0 :	U19=0 :	V19=0
	W19=0 :	X19=0 :	Y19=0 :	Z19=0 : AA19=0 :	AB19=0 :	AC19=0 : 	AD19=0 :	AE19=0

	C40=0 : D40=0 : E40=0 : F40=0 :	G40=0 :	H40=0 :	I40=0 :	J40=0 :	K40=0 :	L40=0 :	M40=0 :	N40=0 :	O40=0 :	P40=0 :	Q40=0 :	R40=0 :	S40=0 :	T40=0 :	U40=0 :	V40=0
	W40=0 :	X40=0 :	Y40=0 :	Z40=0 : AA40=0 :	AB40=0 :	AC40=0 : 	AD40=0 :	AE40=0
	C41=0 : D41=0 : E41=0 : F41=0 :	G41=0 :	H41=0 :	I41=0 :	J41=0 :	K41=0 :	L41=0 :	M41=0 :	N41=0 :	O41=0 :	P41=0 :	Q41=0 :	R41=0 :	S41=0 :	T41=0 :	U41=0 :	V41=0
	W41=0 :	X41=0 :	Y41=0 :	Z41=0 : AA41=0 :	AB41=0 :	AC41=0 : 	AD41=0 :	AE41=0

	C62=0 : D62=0 : E62=0 : F62=0 :	G62=0 :	H62=0 :	I62=0 :	J62=0 :	K62=0 :	L62=0 :	M62=0 :	N62=0 :	O62=0 :	P62=0 :	Q62=0 :	R62=0 :	S62=0
	C63=0 : D63=0 : E63=0 : F63=0 :	G63=0 :	H63=0 :	I63=0 :	J63=0 :	K63=0 :	L63=0 :	M63=0 :	N63=0 :	O63=0 :	P63=0 :	Q63=0 :	R63=0 :	S63=0

While Not rs2.eof 

	Total=0
	BillTypeID=Trim(rs2("BillTypeID"))
	rule1=Trim(rs2("rule1"))
	Total=cdbl(rs2("Total"))

	If BillTypeID="2" Then 
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E18=E18+Total
		End If

		If Mid(rule1,1,5)="12102" Then F18=F18+Total
		If Mid(rule1,1,5)="12109" Then G18=G18+Total
		If Mid(rule1,1,5)="12110" Then H18=H18+Total
		If Mid(rule1,1,5)="13001" Then I18=I18+Total
		If Mid(rule1,1,5)="14001" Then J18=J18+Total
		If Mid(rule1,1,5)="16101" Then K18=K18+Total
		If Mid(rule1,1,5)="16102" Then L18=L18+Total
		If Mid(rule1,1,3)="181" Then M18=M18+Total
		If Mid(rule1,1,5)="21101" Then N18=N18+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O18=O18+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P18=P18+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q18=Q18+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R18=R18+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S18=S18+Total		
		If Mid(rule1,1,2)="22" Then T18=T18+Total		
		If Mid(rule1,1,2)="29" Then U18=U18+Total		
		If left(rule1,2)="29" And Right(Mid(rule1,1,8),1)="1" Then V18=V18+Total	
		If left(rule1,3)="293" And Right(Mid(rule1,1,8),1)="2" Then W18=W18+Total	
		If left(rule1,3)="294" And Right(Mid(rule1,1,8),1)="2" Then X18=X18+Total	
		If Mid(rule1,1,2)="30" Then Y18=Y18+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z18=Z18+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then  AA18=AA18+Total
		If Mid(rule1,1,3)="313" Then AB18=AB18+Total
		If Mid(rule1,1,3)="314" Then AC18=AC18+Total
		If Mid(rule1,1,3)="315" Then AD18=AD18+Total
		If Mid(rule1,1,3)="316" Then AE18=AE18+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C40=C40+Total
		If Mid(rule1,1,3)="321" Then D40=D40+Total
		If left(rule1,2)="32" And Right(Mid(rule1,1,8),1)="1" Then E40=E40+Total	
		If Mid(rule1,1,2)="33" Then F40=F40+Total		
		If Mid(rule1,1,3)="351" Then G40=G40+Total	
		If Mid(rule1,1,3)="352" Then H40=H40+Total	
		If Mid(rule1,1,3)="353" Then I40=I40+Total	
		If Mid(rule1,1,3)="354" Then J40=J40+Total	
		If Mid(rule1,1,3)="361" Then K40=K40+Total	
		If Mid(rule1,1,3)="363" Then L40=L40+Total	
		If Mid(rule1,1,3)="365" Then M40=M40+Total	
		If Mid(rule1,1,3)="381" Then N40=N40+Total	
		If Mid(rule1,1,3)="382" Then O40=O40+Total	
		If Mid(rule1,1,2)="40" Then P40=P40+Total	
		If Mid(rule1,1,5)="43101" Then Q40=Q40+Total	
		If Mid(rule1,1,5)="43102" Then R40=R40+Total	
		If Mid(rule1,1,5)="43103" Then S40=S40+Total	
		If Mid(rule1,1,3)="433" Then T40=T40+Total	
		If Mid(rule1,1,3)="434" Then U40=U40+Total	
		If Mid(rule1,1,3)="441" Then V40=V40+Total	
		If Mid(rule1,1,3)="442" Then W40=W40+Total	
		If Mid(rule1,1,5)="45001" Then X40=X40+Total	
		If Mid(rule1,1,5)="45004" Then Y40=Y40+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z40=Z40+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA40=AA40+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB40=AB40+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC40=AC40+Total			
		If Mid(rule1,1,5)="48105" Then AD40=AD40+Total			
		If Mid(rule1,1,5)="48107" Then AE40=AE40+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C62=C62+Total		
		If Mid(rule1,1,3)="531" Then D62=D62+Total		
		If Mid(rule1,1,3)="532" Then E62=E62+Total		
		If Mid(rule1,1,2)="54" Then F62=F62+Total		
		If Mid(rule1,1,2)="55" Then G62=G62+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H62=H62+Total
		End If
		If Mid(rule1,1,5)="56109" Then I62=I62+Total				
		If Mid(rule1,1,5)="56110" Then J62=J62+Total				
		If Mid(rule1,1,2)="57" Then K62=K62+Total
		If Mid(rule1,1,5)="58003" Then L62=L62+Total
		If Mid(rule1,1,3)="601" Then M62=M62+Total
		If Mid(rule1,1,5)="60203" Then N62=N62+Total
		If Mid(rule1,1,3)="613" Then O62=O62+Total
		If Mid(rule1,1,3)="621" Then P62=P62+Total
		If Mid(rule1,1,3)="622" Then Q62=Q62+Total
		If Mid(rule1,1,3)="623" Then R62=R62+Total
		If Mid(rule1,1,3)="624" Then S62=S62+Total

'---------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E19=E19+Total
		End If
		If Mid(rule1,1,5)="12102" Then F19=F19+Total
		If Mid(rule1,1,5)="12109" Then G19=G19+Total
		If Mid(rule1,1,5)="12110" Then H19=H19+Total
		If Mid(rule1,1,5)="13001" Then I19=I19+Total
		If Mid(rule1,1,5)="14001" Then J19=J19+Total
		If Mid(rule1,1,5)="16101" Then K19=K19+Total
		If Mid(rule1,1,5)="16102" Then L19=L19+Total
		If Mid(rule1,1,3)="181" Then M19=M19+Total
		If Mid(rule1,1,5)="21101" Then N19=N19+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O19=O19+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P19=P19+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q19=Q19+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R19=R19+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S19=S19+Total		
		If Mid(rule1,1,2)="22" Then T19=T19+Total		
		If Mid(rule1,1,2)="29" Then U19=U19+Total		
		If mid(rule1,1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V19=V19+Total	

		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W19=W19+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X19=X19+Total	
		If Mid(rule1,1,2)="30" Then Y19=Y19+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z19=Z19+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA19=AA19+Total
		If Mid(rule1,1,3)="313" Then AB19=AB19+Total
		If Mid(rule1,1,3)="314" Then AC19=AC19+Total
		If Mid(rule1,1,3)="315" Then AD19=AD19+Total
		If Mid(rule1,1,3)="316" Then AE19=AE19+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C41=C41+Total
		If Mid(rule1,1,3)="321" Then D41=D41+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then  E41=E41+Total	
		If Mid(rule1,1,2)="33" Then F41=F41+Total		
		If Mid(rule1,1,3)="351" Then G41=G41+Total	
		If Mid(rule1,1,3)="352" Then H41=H41+Total	
		If Mid(rule1,1,3)="353" Then I41=I41+Total	
		If Mid(rule1,1,3)="354" Then J41=J41+Total	
		If Mid(rule1,1,3)="361" Then K41=K41+Total	
		If Mid(rule1,1,3)="363" Then L41=L41+Total	
		If Mid(rule1,1,3)="365" Then M41=M41+Total	
		If Mid(rule1,1,3)="381" Then N41=N41+Total	
		If Mid(rule1,1,3)="382" Then O41=O41+Total	
		If Mid(rule1,1,2)="40" Then P41=P41+Total	
		If Mid(rule1,1,5)="43101" Then Q41=Q41+Total	
		If Mid(rule1,1,5)="43102" Then R41=R41+Total	
		If Mid(rule1,1,5)="43103" Then S41=S41+Total	
		If Mid(rule1,1,3)="433" Then T41=T41+Total	
		If Mid(rule1,1,3)="434" Then U41=U41+Total	
		If Mid(rule1,1,3)="441" Then V41=V41+Total	
		If Mid(rule1,1,3)="442" Then W41=W41+Total	
		If Mid(rule1,1,5)="45001" Then X41=X41+Total	
		If Mid(rule1,1,5)="45004" Then Y41=Y41+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z41=Z41+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA41=AA41+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB41=AB41+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC41=AC41+Total			
		If Mid(rule1,1,5)="48105" Then AD41=AD41+Total			
		If Mid(rule1,1,5)="48107" Then AE41=AE41+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C63=C63+Total		
		If Mid(rule1,1,3)="531" Then D63=D63+Total		
		If Mid(rule1,1,3)="532" Then E63=E63+Total		
		If Mid(rule1,1,2)="54" Then F63=F63+Total		
		If Mid(rule1,1,2)="55" Then G63=G63+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H63=H63+Total
		End If
		If Mid(rule1,1,5)="56109" Then I63=I63+Total				
		If Mid(rule1,1,5)="56110" Then J63=J63+Total				
		If Mid(rule1,1,2)="57" Then K63=K63+Total
		If Mid(rule1,1,5)="58003" Then L63=L63+Total
		If Mid(rule1,1,3)="601" Then M63=M63+Total
		If Mid(rule1,1,5)="60203" Then N63=N63+Total
		If Mid(rule1,1,3)="613" Then O63=O63+Total
		If Mid(rule1,1,3)="621" Then P63=P63+Total
		If Mid(rule1,1,3)="622" Then Q63=Q63+Total
		If Mid(rule1,1,3)="623" Then R63=R63+Total
		If Mid(rule1,1,3)="624" Then S63=S63+Total
    End if




	rs2.movenext
Wend



strSQL="select BillTypeID,rule1,Total from (Select BillTypeID,rule1,count(billno) as Total from BillBase "
strSQL=strSQL&" where  CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule1"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule2 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
strSQL=strSQL&" and rule2 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule2"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule3 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
strSQL=strSQL&" and rule3 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&"  group by BillTypeID,rule3"
strSQL=strSQL&"  )"
strSQL=strSQL&"  order by 2"
'response.write strsql
'response.end
set rs2=conn.execute(strSQL)
'機車---------------------------------------------------------------------------------------------------------------------------------------------------
	E24=0 : F24=0 :	G24=0 :	H24=0 :	I24=0 :	J24=0 :	K24=0 :	L24=0 :	M24=0 :	N24=0 :	O24=0 :	P24=0 :	Q24=0 :	R24=0 :	S24=0 :	T24=0 :	U24=0 :	V24=0
	W24=0 :	X24=0 :	Y24=0 :	Z24=0 : AA24=0 :	AB24=0 :	AC24=0 : 	AD24=0 :	AE24=0
	E25=0 : F25=0 :	G25=0 :	H25=0 :	I25=0 :	J25=0 :	K25=0 :	L25=0 :	M25=0 :	N25=0 :	O25=0 :	P25=0 :	Q25=0 :	R25=0 :	S25=0 :	T25=0 :	U25=0 :	V25=0
	W25=0 :	X25=0 :	Y25=0 :	Z25=0 : AA25=0 :	AB25=0 :	AC25=0 : 	AD25=0 :	AE25=0

	C46=0 : D46=0 : E46=0 : F46=0 :	G46=0 :	H46=0 :	I46=0 :	J46=0 :	K46=0 :	L46=0 :	M46=0 :	N46=0 :	O46=0 :	P46=0 :	Q46=0 :	R46=0 :	S46=0 :	T46=0 :	U46=0 :	V46=0
	W46=0 :	X46=0 :	Y46=0 :	Z46=0 : AA46=0 :	AB46=0 :	AC46=0 : 	AD46=0 :	AE46=0
	C47=0 : D47=0 : E47=0 : F47=0 :	G47=0 :	H47=0 :	I47=0 :	J47=0 :	K47=0 :	L47=0 :	M47=0 :	N47=0 :	O47=0 :	P47=0 :	Q47=0 :	R47=0 :	S47=0 :	T47=0 :	U47=0 :	V47=0
	W47=0 :	X47=0 :	Y47=0 :	Z47=0 : AA47=0 :	AB47=0 :	AC47=0 : 	AD47=0 :	AE47=0

	C68=0 : D68=0 : E68=0 : F68=0 :	G68=0 :	H68=0 :	I68=0 :	J68=0 :	K68=0 :	L68=0 :	M68=0 :	N68=0 :	O68=0 :	P68=0 :	Q68=0 :	R68=0 :	S68=0
	C69=0 : D69=0 : E69=0 : F69=0 :	G69=0 :	H69=0 :	I69=0 :	J69=0 :	K69=0 :	L69=0 :	M69=0 :	N69=0 :	O69=0 :	P69=0 :	Q69=0 :	R69=0 :	S69=0

While Not rs2.eof 

	Total=0
	BillTypeID=Trim(rs2("BillTypeID"))
	rule1=Trim(rs2("rule1"))
	Total=cdbl(rs2("Total"))

	If BillTypeID="2" Then 
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E24=E24+Total
		End If

		If Mid(rule1,1,5)="12102" Then F24=F24+Total
		If Mid(rule1,1,5)="12109" Then G24=G24+Total
		If Mid(rule1,1,5)="12110" Then H24=H24+Total
		If Mid(rule1,1,5)="13001" Then I24=I24+Total
		If Mid(rule1,1,5)="14001" Then J24=J24+Total
		If Mid(rule1,1,5)="16101" Then K24=K24+Total
		If Mid(rule1,1,5)="16102" Then L24=L24+Total
		If Mid(rule1,1,3)="181" Then M24=M24+Total
		If Mid(rule1,1,5)="21101" Then N24=N24+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O24=O24+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P24=P24+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q24=Q24+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R24=R24+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S24=S24+Total		
		If Mid(rule1,1,2)="22" Then T24=T24+Total		
		If Mid(rule1,1,2)="29" Then U24=U24+Total		
		If left(rule1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V24=V24+Total	
		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W24=W24+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X24=X24+Total	
		If Mid(rule1,1,2)="30" Then Y24=Y24+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z24=Z24+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA24=AA24+Total
		If Mid(rule1,1,3)="313" Then AB24=AB24+Total
		If Mid(rule1,1,3)="314" Then AC24=AC24+Total
		If Mid(rule1,1,3)="315" Then AD24=AD24+Total
		If Mid(rule1,1,3)="316" Then AE24=AE24+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C46=C46+Total
		If Mid(rule1,1,3)="321" Then D46=D46+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then E46=E46+Total	
		If Mid(rule1,1,2)="30" Then F46=F46+Total		
		If Mid(rule1,1,3)="351" Then G46=G46+Total	
		If Mid(rule1,1,3)="352" Then H46=H46+Total	
		If Mid(rule1,1,3)="353" Then I46=I46+Total	
		If Mid(rule1,1,3)="354" Then J46=J46+Total	
		If Mid(rule1,1,3)="361" Then K46=K46+Total	
		If Mid(rule1,1,3)="363" Then L46=L46+Total	
		If Mid(rule1,1,3)="365" Then M46=M46+Total	
		If Mid(rule1,1,3)="381" Then N46=N46+Total	
		If Mid(rule1,1,3)="382" Then O46=O46+Total	
		If Mid(rule1,1,2)="40" Then P46=P46+Total	
		If Mid(rule1,1,5)="43101" Then Q46=Q46+Total	
		If Mid(rule1,1,5)="43102" Then R46=R46+Total	
		If Mid(rule1,1,5)="43103" Then S46=S46+Total	
		If Mid(rule1,1,3)="433" Then T46=T46+Total	
		If Mid(rule1,1,3)="434" Then U46=U46+Total	
		If Mid(rule1,1,3)="441" Then V46=V46+Total	
		If Mid(rule1,1,3)="442" Then W46=W46+Total	
		If Mid(rule1,1,5)="45001" Then X46=X46+Total	
		If Mid(rule1,1,5)="45004" Then Y46=Y46+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z46=Z46+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA46=AA46+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB46=AB46+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC46=AC46+Total			
		If Mid(rule1,1,5)="48105" Then AD46=AD46+Total			
		If Mid(rule1,1,5)="48107" Then AE46=AE46+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C68=C68+Total		
		If Mid(rule1,1,3)="531" Then D68=D68+Total		
		If Mid(rule1,1,3)="532" Then E68=E68+Total		
		If Mid(rule1,1,2)="54" Then F68=F68+Total		
		If Mid(rule1,1,2)="55" Then G68=G68+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H68=H68+Total
		End If
		If Mid(rule1,1,5)="56109" Then I68=I68+Total				
		If Mid(rule1,1,5)="56110" Then J68=J68+Total				
		If Mid(rule1,1,2)="57" Then K68=K68+Total
		If Mid(rule1,1,5)="58003" Then L68=L68+Total
		If Mid(rule1,1,3)="601" Then M68=M68+Total
		If Mid(rule1,1,5)="60203" Then N68=N68+Total
		If Mid(rule1,1,3)="613" Then O68=O68+Total
		If Mid(rule1,1,3)="621" Then P68=P68+Total
		If Mid(rule1,1,3)="622" Then Q68=Q68+Total
		If Mid(rule1,1,3)="623" Then R68=R68+Total
		If Mid(rule1,1,3)="624" Then S68=S68+Total

'---------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E25=E25+Total
		End If
		If Mid(rule1,1,5)="12102" Then F25=F25+Total
		If Mid(rule1,1,5)="12109" Then G25=G25+Total
		If Mid(rule1,1,5)="12110" Then H25=H25+Total
		If Mid(rule1,1,5)="13001" Then I25=I25+Total
		If Mid(rule1,1,5)="14001" Then J25=J25+Total
		If Mid(rule1,1,5)="16101" Then K25=K25+Total
		If Mid(rule1,1,5)="16102" Then L25=L25+Total
		If Mid(rule1,1,3)="181" Then M25=M25+Total
		If Mid(rule1,1,5)="21101" Then N25=N25+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O25=O25+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P25=P25+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q25=Q25+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R25=R25+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S25=S25+Total		
		If Mid(rule1,1,2)="22" Then T25=T25+Total		
		If Mid(rule1,1,2)="29" Then U25=U25+Total		
		If left(rule1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V25=V25+Total	
		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W25=W25+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X25=X25+Total	
		If Mid(rule1,1,2)="30" Then Y25=Y25+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z25=Z25+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA25=AA25+Total
		If Mid(rule1,1,3)="313" Then AB25=AB25+Total
		If Mid(rule1,1,3)="314" Then AC25=AC25+Total
		If Mid(rule1,1,3)="315" Then AD25=AD25+Total
		If Mid(rule1,1,3)="316" Then AE25=AE25+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C47=C47+Total
		If Mid(rule1,1,3)="321" Then D47=D47+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then E47=E47+Total	
		If Mid(rule1,1,2)="30" Then F47=F47+Total		
		If Mid(rule1,1,3)="351" Then G47=G47+Total	
		If Mid(rule1,1,3)="352" Then H47=H47+Total	
		If Mid(rule1,1,3)="353" Then I47=I47+Total	
		If Mid(rule1,1,3)="354" Then J47=J47+Total	
		If Mid(rule1,1,3)="361" Then K47=K47+Total	
		If Mid(rule1,1,3)="363" Then L47=L47+Total	
		If Mid(rule1,1,3)="365" Then M47=M47+Total	
		If Mid(rule1,1,3)="381" Then N47=N47+Total	
		If Mid(rule1,1,3)="382" Then O47=O47+Total	
		If Mid(rule1,1,2)="40" Then P47=P47+Total	
		If Mid(rule1,1,5)="43101" Then Q47=Q47+Total	
		If Mid(rule1,1,5)="43102" Then R47=R47+Total	
		If Mid(rule1,1,5)="43103" Then S47=S47+Total	
		If Mid(rule1,1,3)="433" Then T47=T47+Total	
		If Mid(rule1,1,3)="434" Then U47=U47+Total	
		If Mid(rule1,1,3)="441" Then V47=V47+Total	
		If Mid(rule1,1,3)="442" Then W47=W47+Total	
		If Mid(rule1,1,5)="45001" Then X47=X47+Total	
		If Mid(rule1,1,5)="45004" Then Y47=Y47+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z47=Z47+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA47=AA47+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB47=AB47+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC47=AC47+Total			
		If Mid(rule1,1,5)="48105" Then AD47=AD47+Total			
		If Mid(rule1,1,5)="48107" Then AE47=AE47+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C69=C69+Total		
		If Mid(rule1,1,3)="531" Then D69=D69+Total		
		If Mid(rule1,1,3)="532" Then E69=E69+Total		
		If Mid(rule1,1,2)="54" Then F69=F69+Total		
		If Mid(rule1,1,2)="55" Then G69=G69+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H69=H69+Total
		End If
		If Mid(rule1,1,5)="56109" Then I69=I69+Total				
		If Mid(rule1,1,5)="56110" Then J69=J69+Total				
		If Mid(rule1,1,2)="57" Then K69=K69+Total
		If Mid(rule1,1,5)="58003" Then L69=L69+Total
		If Mid(rule1,1,3)="601" Then M69=M69+Total
		If Mid(rule1,1,5)="60203" Then N69=N69+Total
		If Mid(rule1,1,3)="613" Then O69=O69+Total
		If Mid(rule1,1,3)="621" Then P69=P69+Total
		If Mid(rule1,1,3)="622" Then Q69=Q69+Total
		If Mid(rule1,1,3)="623" Then R69=R69+Total
		If Mid(rule1,1,3)="624" Then S69=S69+Total
    End if




	rs2.movenext
wend
'--------------------------------------------------------------------------------------------------------------------------------------------------------------

strSQL="select BillTypeID,rule1,Total from (Select BillTypeID,rule1,count(billno) as Total from BillBase "
strSQL=strSQL&" where  CarAddID='9'"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule1"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule2 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarAddID='9'"
strSQL=strSQL&" and rule2 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule2"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule3 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarAddID='9'"
strSQL=strSQL&" and rule3 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&"  group by BillTypeID,rule3"
strSQL=strSQL&"  )"
strSQL=strSQL&"  order by 2"
'response.write strsql
'response.end
set rs2=conn.execute(strSQL)
'550重機車---------------------------------------------------------------------------------------------------------------------------------------------------
	E21=0 : F21=0 :	G21=0 :	H21=0 :	I21=0 :	J21=0 :	K21=0 :	L21=0 :	M21=0 :	N21=0 :	O21=0 :	P21=0 :	Q21=0 :	R21=0 :	S21=0 :	T21=0 :	U21=0 :	V21=0
	W21=0 :	X21=0 :	Y21=0 :	Z21=0 : AA21=0 :	AB21=0 :	AC21=0 : 	AD21=0 :	AE21=0
	E22=0 : F22=0 :	G22=0 :	H22=0 :	I22=0 :	J22=0 :	K22=0 :	L22=0 :	M22=0 :	N22=0 :	O22=0 :	P22=0 :	Q22=0 :	R22=0 :	S22=0 :	T22=0 :	U22=0 :	V22=0
	W22=0 :	X22=0 :	Y22=0 :	Z22=0 : AA22=0 :	AB22=0 :	AC22=0 : 	AD22=0 :	AE22=0

	C43=0 : D43=0 : E43=0 : F43=0 :	G43=0 :	H43=0 :	I43=0 :	J43=0 :	K43=0 :	L43=0 :	M43=0 :	N43=0 :	O43=0 :	P43=0 :	Q43=0 :	R43=0 :	S43=0 :	T43=0 :	U43=0 :	V43=0
	W43=0 :	X43=0 :	Y43=0 :	Z43=0 : AA43=0 :	AB43=0 :	AC43=0 : 	AD43=0 :	AE43=0
	C44=0 : D44=0 : E44=0 : F44=0 :	G44=0 :	H44=0 :	I44=0 :	J44=0 :	K44=0 :	L44=0 :	M44=0 :	N44=0 :	O44=0 :	P44=0 :	Q44=0 :	R44=0 :	S44=0 :	T44=0 :	U44=0 :	V44=0
	W44=0 :	X44=0 :	Y44=0 :	Z44=0 : AA44=0 :	AB44=0 :	AC44=0 : 	AD44=0 :	AE44=0

	C65=0 : D65=0 : E65=0 : F65=0 :	G65=0 :	H65=0 :	I65=0 :	J65=0 :	K65=0 :	L65=0 :	M65=0 :	N65=0 :	O65=0 :	P65=0 :	Q65=0 :	R65=0 :	S65=0
	C66=0 : D66=0 : E66=0 : F66=0 :	G66=0 :	H66=0 :	I66=0 :	J66=0 :	K66=0 :	L66=0 :	M66=0 :	N66=0 :	O66=0 :	P66=0 :	Q66=0 :	R66=0 :	S66=0

While Not rs2.eof 

	Total=0
	BillTypeID=Trim(rs2("BillTypeID"))
	rule1=Trim(rs2("rule1"))
	Total=cdbl(rs2("Total"))

	If BillTypeID="2" Then 
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E21=E21+Total
		End If

		If Mid(rule1,1,5)="12102" Then F21=F21+Total
		If Mid(rule1,1,5)="12109" Then G21=G21+Total
		If Mid(rule1,1,5)="12110" Then H21=H21+Total
		If Mid(rule1,1,5)="13001" Then I21=I21+Total
		If Mid(rule1,1,5)="14001" Then J21=J21+Total
		If Mid(rule1,1,5)="16101" Then K21=K21+Total
		If Mid(rule1,1,5)="16102" Then L21=L21+Total
		If Mid(rule1,1,3)="181" Then M21=M21+Total
		If Mid(rule1,1,5)="21101" Then N21=N21+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O21=O21+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P21=P21+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q21=Q21+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R21=R21+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S21=S21+Total		
		If Mid(rule1,1,2)="22" Then T21=T21+Total		
		If Mid(rule1,1,2)="29" Then U21=U21+Total		
		If left(rule1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V21=V21+Total	
		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W21=W21+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X21=X21+Total	
		If Mid(rule1,1,2)="30" Then Y21=Y21+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z21=Z21+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA21=AA21+Total
		If Mid(rule1,1,3)="313" Then AB21=AB21+Total
		If Mid(rule1,1,3)="314" Then AC21=AC21+Total
		If Mid(rule1,1,3)="315" Then AD21=AD21+Total
		If Mid(rule1,1,3)="316" Then AE21=AE21+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C43=C43+Total
		If Mid(rule1,1,3)="321" Then D43=D43+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then E43=E43+Total	
		If Mid(rule1,1,2)="30" Then F43=F43+Total		
		If Mid(rule1,1,3)="351" Then G43=G43+Total	
		If Mid(rule1,1,3)="352" Then H43=H43+Total	
		If Mid(rule1,1,3)="353" Then I43=I43+Total	
		If Mid(rule1,1,3)="354" Then J43=J43+Total	
		If Mid(rule1,1,3)="361" Then K43=K43+Total	
		If Mid(rule1,1,3)="363" Then L43=L43+Total	
		If Mid(rule1,1,3)="365" Then M43=M43+Total	
		If Mid(rule1,1,3)="381" Then N43=N43+Total	
		If Mid(rule1,1,3)="382" Then O43=O43+Total	
		If Mid(rule1,1,2)="40" Then P43=P43+Total	
		If Mid(rule1,1,5)="43101" Then Q43=Q43+Total	
		If Mid(rule1,1,5)="43102" Then R43=R43+Total	
		If Mid(rule1,1,5)="43103" Then S43=S43+Total	
		If Mid(rule1,1,3)="433" Then T43=T43+Total	
		If Mid(rule1,1,3)="434" Then U43=U43+Total	
		If Mid(rule1,1,3)="441" Then V43=V43+Total	
		If Mid(rule1,1,3)="442" Then W43=W43+Total	
		If Mid(rule1,1,5)="45001" Then X43=X43+Total	
		If Mid(rule1,1,5)="45004" Then Y43=Y43+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z43=Z43+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA43=AA43+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB43=AB43+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC43=AC43+Total			
		If Mid(rule1,1,5)="48105" Then AD43=AD43+Total			
		If Mid(rule1,1,5)="48107" Then AE43=AE43+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C65=C65+Total		
		If Mid(rule1,1,3)="531" Then D65=D65+Total		
		If Mid(rule1,1,3)="532" Then E65=E65+Total		
		If Mid(rule1,1,2)="54" Then F65=F65+Total		
		If Mid(rule1,1,2)="55" Then G65=G65+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H65=H65+Total
		End If
		If Mid(rule1,1,5)="56109" Then I65=I65+Total				
		If Mid(rule1,1,5)="56110" Then J65=J65+Total				
		If Mid(rule1,1,2)="57" Then K65=K65+Total
		If Mid(rule1,1,5)="58003" Then L65=L65+Total
		If Mid(rule1,1,3)="601" Then M65=M65+Total
		If Mid(rule1,1,5)="60203" Then N65=N65+Total
		If Mid(rule1,1,3)="613" Then O65=O65+Total
		If Mid(rule1,1,3)="621" Then P65=P65+Total
		If Mid(rule1,1,3)="622" Then Q65=Q65+Total
		If Mid(rule1,1,3)="623" Then R65=R65+Total
		If Mid(rule1,1,3)="624" Then S65=S65+Total

'---------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------
    Else
		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E22=E22+Total
		End If
		If Mid(rule1,1,5)="12102" Then F22=F22+Total
		If Mid(rule1,1,5)="12109" Then G22=G22+Total
		If Mid(rule1,1,5)="12110" Then H22=H22+Total
		If Mid(rule1,1,5)="13001" Then I22=I22+Total
		If Mid(rule1,1,5)="14001" Then J22=J22+Total
		If Mid(rule1,1,5)="16101" Then K22=K22+Total
		If Mid(rule1,1,5)="16102" Then L22=L22+Total
		If Mid(rule1,1,3)="181" Then M22=M22+Total
		If Mid(rule1,1,5)="21101" Then N22=N22+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O22=O22+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P22=P22+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q22=Q22+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R22=R22+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S22=S22+Total		
		If Mid(rule1,1,2)="22" Then T22=T22+Total		
		If Mid(rule1,1,2)="29" Then U22=U22+Total		
		If left(rule1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V22=V22+Total	
		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W22=W22+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X22=X22+Total	
		If Mid(rule1,1,2)="30" Then Y22=Y22+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z22=Z22+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA22=AA22+Total
		If Mid(rule1,1,3)="313" Then AB22=AB22+Total
		If Mid(rule1,1,3)="314" Then AC22=AC22+Total
		If Mid(rule1,1,3)="315" Then AD22=AD22+Total
		If Mid(rule1,1,3)="316" Then AE22=AE22+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C44=C44+Total
		If Mid(rule1,1,3)="321" Then D44=D44+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then E44=E44+Total	
		If Mid(rule1,1,2)="30" Then F44=F44+Total		
		If Mid(rule1,1,3)="351" Then G44=G44+Total	
		If Mid(rule1,1,3)="352" Then H44=H44+Total	
		If Mid(rule1,1,3)="353" Then I44=I44+Total	
		If Mid(rule1,1,3)="354" Then J44=J44+Total	
		If Mid(rule1,1,3)="361" Then K44=K44+Total	
		If Mid(rule1,1,3)="363" Then L44=L44+Total	
		If Mid(rule1,1,3)="365" Then M44=M44+Total	
		If Mid(rule1,1,3)="381" Then N44=N44+Total	
		If Mid(rule1,1,3)="382" Then O44=O44+Total	
		If Mid(rule1,1,2)="40" Then P44=P44+Total	
		If Mid(rule1,1,5)="43101" Then Q44=Q44+Total	
		If Mid(rule1,1,5)="43102" Then R44=R44+Total	
		If Mid(rule1,1,5)="43103" Then S44=S44+Total	
		If Mid(rule1,1,3)="433" Then T44=T44+Total	
		If Mid(rule1,1,3)="434" Then U44=U44+Total	
		If Mid(rule1,1,3)="441" Then V44=V44+Total	
		If Mid(rule1,1,3)="442" Then W44=W44+Total	
		If Mid(rule1,1,5)="45001" Then X44=X44+Total	
		If Mid(rule1,1,5)="45004" Then Y44=Y44+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z44=Z44+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA44=AA44+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB44=AB44+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC44=AC44+Total			
		If Mid(rule1,1,5)="48105" Then AD44=AD44+Total			
		If Mid(rule1,1,5)="48107" Then AE44=AE44+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C66=C66+Total		
		If Mid(rule1,1,3)="531" Then D66=D66+Total		
		If Mid(rule1,1,3)="532" Then E66=E66+Total		
		If Mid(rule1,1,2)="54" Then F66=F66+Total		
		If Mid(rule1,1,2)="55" Then G66=G66+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H66=H66+Total
		End If
		If Mid(rule1,1,5)="56109" Then I66=I66+Total				
		If Mid(rule1,1,5)="56110" Then J66=J66+Total				
		If Mid(rule1,1,2)="57" Then K66=K66+Total
		If Mid(rule1,1,5)="58003" Then L66=L66+Total
		If Mid(rule1,1,3)="601" Then M66=M66+Total
		If Mid(rule1,1,5)="60203" Then N66=N66+Total
		If Mid(rule1,1,3)="613" Then O66=O66+Total
		If Mid(rule1,1,3)="621" Then P66=P66+Total
		If Mid(rule1,1,3)="622" Then Q66=Q66+Total
		If Mid(rule1,1,3)="623" Then R66=R66+Total
		If Mid(rule1,1,3)="624" Then S66=S66+Total
    End if




	rs2.movenext
wend



'--------------------------------------------------------------------------------------------------------------------------------------------------------------

strSQL="select BillTypeID,rule1,Total from (Select BillTypeID,rule1,count(billno) as Total from BillBase "
strSQL=strSQL&" where  CarAddID='5'"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule1"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule2 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarAddID='5'"
strSQL=strSQL&" and rule2 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&" group by BillTypeID,rule2"
strSQL=strSQL&" union all"
strSQL=strSQL&" Select BillTypeID,rule3 as rule1,count(billno) as Total from BillBase  where "
strSQL=strSQL&"  CarAddID='5'"
strSQL=strSQL&" and rule3 is not null"
strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
strSQL=strSQL& tmpSql
strSQL=strSQL&"  group by BillTypeID,rule3"
strSQL=strSQL&"  )"
strSQL=strSQL&"  order by 2"
'response.write strsql
'response.end
set rs2=conn.execute(strSQL)
'動力機械---------------------------------------------------------------------------------------------------------------------------------------------------
	E26=0 : F26=0 :	G26=0 :	H26=0 :	I26=0 :	J26=0 :	K26=0 :	L26=0 :	M26=0 :	N26=0 :	O26=0 :	P26=0 :	Q26=0 :	R26=0 :	S26=0 :	T26=0 :	U26=0 :	V26=0
	W26=0 :	X26=0 :	Y26=0 :	Z26=0 : AA26=0 :	AB26=0 :	AC26=0 : 	AD26=0 :	AE26=0


	C48=0 : D48=0 : E48=0 : F48=0 :	G48=0 :	H48=0 :	I48=0 :	J48=0 :	K48=0 :	L48=0 :	M48=0 :	N48=0 :	O48=0 :	P48=0 :	Q48=0 :	R48=0 :	S48=0 :	T48=0 :	U48=0 :	V48=0
	W48=0 :	X48=0 :	Y48=0 :	Z48=0 : AA48=0 :	AB48=0 :	AC48=0 : 	AD48=0 :	AE48=0

	C70=0 : D70=0 : E70=0 : F70=0 :	G70=0 :	H70=0 :	I70=0 :	J70=0 :	K70=0 :	L70=0 :	M70=0 :	N70=0 :	O70=0 :	P70=0 :	Q70=0 :	R70=0 :	S70=0


While Not rs2.eof 

	Total=0
	BillTypeID=Trim(rs2("BillTypeID"))
	rule1=Trim(rs2("rule1"))
	Total=cdbl(rs2("Total"))

		If mid(rule1,1,5)="12101" Or mid(rule1,1,5)="12103" Or mid(rule1,1,5)="12104" Or mid(rule1,1,5)="12105" Or mid(rule1,1,5)="12106" Or mid(rule1,1,5)="12107" Or mid(rule1,1,5)="12108" Then 
			E21=E21+Total
		End If

		If Mid(rule1,1,5)="12102" Then F26=F26+Total
		If Mid(rule1,1,5)="12109" Then G26=G26+Total
		If Mid(rule1,1,5)="12110" Then H26=H26+Total
		If Mid(rule1,1,5)="13001" Then I26=I26+Total
		If Mid(rule1,1,5)="14001" Then J26=J26+Total
		If Mid(rule1,1,5)="16101" Then K26=K26+Total
		If Mid(rule1,1,5)="16102" Then L26=L26+Total
		If Mid(rule1,1,3)="181" Then M26=M26+Total
		If Mid(rule1,1,5)="21101" Then N26=N26+Total
		If Mid(rule1,1,7)="2110201" Or Mid(rule1,1,7)="2110202" Or Mid(rule1,1,7)="2110301" Or Mid(rule1,1,7)="2110302" Or Mid(rule1,1,7)="2110401" Or Mid(rule1,1,7)="2110402" Or Mid(rule1,1,7)="2110601" Or Mid(rule1,1,7)="2110602" Or Mid(rule1,1,7)="2110701" Or Mid(rule1,1,7)="2110702" Or Mid(rule1,1,7)="2110801" Or Mid(rule1,1,7)="2110802" Or Mid(rule1,1,7)="2110901" Or Mid(rule1,1,7)="2110902" Then 
			O26=O26+Total
		End If
		If Mid(rule1,1,7)="2110501" Then P26=P26+Total		
		If Mid(rule1,1,8)="21101011" Or Mid(rule1,1,8)="21101021" Then Q26=Q26+Total		
		If Mid(rule1,1,8)="21102011" Or Mid(rule1,1,8)="21103021" Or Mid(rule1,1,8)="21103011" Or Mid(rule1,1,8)="21104011" Or Mid(rule1,1,8)="21105011" Or Mid(rule1,1,8)="21105021" Or Mid(rule1,1,8)="21105031" Or Mid(rule1,1,8)="21105041" Or Mid(rule1,1,8)="21106011" Or Mid(rule1,1,8)="21106021" Or Mid(rule1,1,8)="21106031" Or Mid(rule1,1,8)="21106041" Or Mid(rule1,1,8)="21106051" Or Mid(rule1,1,8)="21106011" Then 
			R26=R26+Total		
		End If
		If Mid(rule1,1,8)="21107011" Then S26=S26+Total		
		If Mid(rule1,1,2)="22" Then T26=T26+Total		
		If Mid(rule1,1,2)="29" Then U26=U26+Total		
		If left(rule1,2)="29" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then V26=V26+Total	
		If left(rule1,3)="293" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then W26=W26+Total	
		If left(rule1,3)="294" Then If len(rule1)=8 Then If  Right(rule1,1)="2" Then X26=X26+Total	
		If Mid(rule1,1,2)="30" Then Y26=Y26+Total		
		If Mid(rule1,1,7)="3110001" Or Mid(rule1,1,7)="3110002" Or Mid(rule1,1,7)="3110003" Or Mid(rule1,1,7)="3110004" Then 
			Z26=Z26+Total		
		End If
		If rule1="3120002" Or rule1="3120001" Then AA26=AA26+Total
		If Mid(rule1,1,3)="313" Then AB26=AB26+Total
		If Mid(rule1,1,3)="314" Then AC26=AC26+Total
		If Mid(rule1,1,3)="315" Then AD26=AD26+Total
		If Mid(rule1,1,3)="316" Then AE26=AE26+Total
'----------------------------------------------------------------------------------------------------------------------------------------------------
		If Mid(rule1,1,8)="31100011" Or Mid(rule1,1,8)="31100021" Or Mid(rule1,1,8)="31200011" Or Mid(rule1,1,8)="31200021" Then C48=C48+Total
		If Mid(rule1,1,3)="321" Then D48=D48+Total
		If left(rule1,2)="32" Then If len(rule1)=8 Then If  Right(rule1,1)="1" Then E48=E48+Total	
		If Mid(rule1,1,2)="30" Then F48=F48+Total		
		If Mid(rule1,1,3)="351" Then G48=G48+Total	
		If Mid(rule1,1,3)="352" Then H48=H48+Total	
		If Mid(rule1,1,3)="353" Then I48=I48+Total	
		If Mid(rule1,1,3)="354" Then J48=J48+Total	
		If Mid(rule1,1,3)="361" Then K48=K48+Total	
		If Mid(rule1,1,3)="363" Then L48=L48+Total	
		If Mid(rule1,1,3)="365" Then M48=M48+Total	
		If Mid(rule1,1,3)="381" Then N48=N48+Total	
		If Mid(rule1,1,3)="382" Then O48=O48+Total	
		If Mid(rule1,1,2)="40" Then P48=P48+Total	
		If Mid(rule1,1,5)="43101" Then Q48=Q48+Total	
		If Mid(rule1,1,5)="43102" Then R48=R48+Total	
		If Mid(rule1,1,5)="43103" Then S48=S48+Total	
		If Mid(rule1,1,3)="433" Then T48=T48+Total	
		If Mid(rule1,1,3)="434" Then U48=U48+Total	
		If Mid(rule1,1,3)="441" Then V48=V48+Total	
		If Mid(rule1,1,3)="442" Then W48=W48+Total	
		If Mid(rule1,1,5)="45001" Then X48=X48+Total	
		If Mid(rule1,1,5)="45004" Then Y48=Y48+Total	
		If Mid(rule1,1,5)="45002" Or Mid(rule1,1,5)="45003" Or Mid(rule1,1,5)="45005" Or Mid(rule1,1,5)="45006" Or Mid(rule1,1,5)="45007" Or Mid(rule1,1,5)="45008" Or Mid(rule1,1,5)="45009" Or Mid(rule1,1,5)="45010" Or Mid(rule1,1,5)="45011" Or Mid(rule1,1,5)="45012" Or Mid(rule1,1,5)="45013" Or Mid(rule1,1,5)="45014" Or Mid(rule1,1,5)="45015" Then 
			Z48=Z48+Total	
		End If
		If Mid(rule1,1,2)="47" Then AA48=AA48+Total	
		If Mid(rule1,1,5)="48101" Or Mid(rule1,1,5)="48102" Or Mid(rule1,1,5)="48103" Or Mid(rule1,1,5)="48106" Then 
			AB48=AB48+Total	
		End If
		If Mid(rule1,1,5)="48104" Then AC48=AC48+Total			
		If Mid(rule1,1,5)="48105" Then AD48=AD48+Total			
		If Mid(rule1,1,5)="48107" Then AE48=AE48+Total		
'----------------------------------------------------------------------------------------------------------------------------------------------------	
		If Mid(rule1,1,3)="482" Then C70=C70+Total		
		If Mid(rule1,1,3)="531" Then D70=D70+Total		
		If Mid(rule1,1,3)="532" Then E70=E70+Total		
		If Mid(rule1,1,2)="54" Then F70=F70+Total		
		If Mid(rule1,1,2)="55" Then G70=G70+Total		

		If Mid(rule1,1,5)="56101" Or Mid(rule1,1,5)="56102" Or Mid(rule1,1,5)="56103" Or Mid(rule1,1,5)="56104" Or Mid(rule1,1,5)="56105" Or Mid(rule1,1,5)="56106" Or Mid(rule1,1,5)="56107" Or Mid(rule1,1,5)="56108" Then 
			H70=H70+Total
		End If
		If Mid(rule1,1,5)="56109" Then I70=I70+Total				
		If Mid(rule1,1,5)="56110" Then J70=J70+Total				
		If Mid(rule1,1,2)="57" Then K70=K70+Total
		If Mid(rule1,1,5)="58003" Then L70=L70+Total
		If Mid(rule1,1,3)="601" Then M70=M70+Total
		If Mid(rule1,1,5)="60203" Then N70=N70+Total
		If Mid(rule1,1,3)="613" Then O70=O70+Total
		If Mid(rule1,1,3)="621" Then P70=P70+Total
		If Mid(rule1,1,3)="622" Then Q70=Q70+Total
		If Mid(rule1,1,3)="623" Then R70=R70+Total
		If Mid(rule1,1,3)="624" Then S70=S70+Total


	rs2.movenext
wend

function GetValue(CarType,TypeID,RuleData,RuleType)
                          '   車種        攔逕     法條       法條式
           ' BillTypeID TypeID 1攔停 2逕舉
   	 lawList = RuleData
   	 lawList = Split(lawList,",")   	  
	 RuleData=""
   	 For i = 0 To UBound(lawList)
		tmpAry = lawList(i)
		tmpAry = Split(tmpAry,"$")
		if RuleData<>"" then RuleData=trim(RuleData)&"','"
		RuleData=trim(RuleData)&tmpAry(0)
   	 Next

           if CarType=1 then 
            '汽車
             TypeStr=" and CarSimpleID in ('2','1')  and (CarAddID not in ('9','5') or CarAddID is null)"
           elseif CarType=2 then 
            '輕機+一般重機
             TypeStr=" and CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
           elseif CarType=3 then 
            '大型重機
             TypeStr=" and (CarSimpleID='3' and CarNO Like 'FA%')"
           elseif CarType=4 then 
            '550CC重機
             TypeStr=" and CarAddID='9'"
           elseif CarType=5 then 
            '動力
             TypeStr=" and CarAddID='5'"
           end if


            strSQL= "select sum(total) as total from (" 

			strSQL=strSQL&"Select count(billno) as Total from BillBase "
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			
			if RuleType=1 then 
			    '條
    			tempRule=" and (Substr(Rule1,1,2) in ('"&RuleData&"') ) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and (Substr(Rule1,1,3) in ('"&RuleData&"') ) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule1,1,5) in ('"&RuleData&"') ) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and (Substr(Rule1,1,7) in ('"&RuleData&"') ) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule1 in ('"&RuleData&"') ) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ((Substr(Rule1,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule1,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql
			strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBase "
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			if RuleType=1 then 
			    '條
    			tempRule=" and ( Substr(Rule2,1,2) in ('"&RuleData&"')) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and ( Substr(Rule2,1,3) in ('"&RuleData&"')) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule2,1,5) in ('"&RuleData&"')) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and ( Substr(Rule2,1,7) in ('"&RuleData&"')) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule2 in ('"&RuleData&"')) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ( (Substr(Rule2,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule2,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql

			strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBase "
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			if RuleType=1 then 
			    '條
    			tempRule=" and ( Substr(Rule3,1,2) in ('"&RuleData&"')) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and ( Substr(Rule3,1,3) in ('"&RuleData&"')) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule3,1,5) in ('"&RuleData&"')) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and ( Substr(Rule3,1,7) in ('"&RuleData&"')) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule3 in ('"&RuleData&"')) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ( (Substr(Rule3,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule3,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql
			strSQL= strSQL& ")" 
'response.write strsql

			set rs2=conn.execute(strSQL)
			GetValue=CDbl(rs2("Total"))
            rs2.close

            set rs2=nothing

end function    

function GetValue2(RuleData,RuleType)

           ' BillTypeID TypeID 1攔停 2逕舉


   	 lawList = RuleData
   	 lawList = Split(lawList,",")   	  
	 RuleData=""
   	 For i = 0 To UBound(lawList)
		tmpAry = lawList(i)
		tmpAry = Split(tmpAry,"$")
		if RuleData<>"" then RuleData=trim(RuleData)&"','"
		RuleData=trim(RuleData)&tmpAry(0)
   	 Next

'	 response.write RuleData
'	 response.end
            strSQL= "select sum(total) as total from (" 
			strSQL=strSQL&"Select count(billno) as Total from BillBaseVIEW where 1=1 " 
			
			if RuleType=1 then 
			    '條
    			tempRule=" and (Substr(Rule1,1,2) in ('"&RuleData&"') ) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and (Substr(Rule1,1,3) in ('"&RuleData&"') ) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule1,1,5) in ('"&RuleData&"') ) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and (Substr(Rule1,1,7) in ('"&RuleData&"') ) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule1 in ('"&RuleData&"') ) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ((Substr(Rule1,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule1,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
            end if            
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql
			strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBaseVIEW where 1=1 " 
			if RuleType=1 then 
			    '條
    			tempRule=" and Substr(Rule2,1,2) in ('"&RuleData&"') "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and  Substr(Rule2,1,3) in ('"&RuleData&"') "
			elseif RuleType=3 then 
                '款
    			tempRule=" and  Substr(Rule2,1,5) in ('"&RuleData&"') "
			elseif RuleType=4 then 
                '號
    			tempRule=" and  Substr(Rule2,1,7) in ('"&RuleData&"') "
			elseif RuleType=5 then 
                '全
    			tempRule=" and  Rule2 in ('"&RuleData&"') "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and  (Substr(Rule2,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) "
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and (Substr(Rule2,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) "
            end if   
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql

		strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBaseVIEW where 1=1 " 
			if RuleType=1 then 
			    '條
    			tempRule=" and Substr(Rule3,1,2) in ('"&RuleData&"') "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and  Substr(Rule3,1,3) in ('"&RuleData&"') "
			elseif RuleType=3 then 
                '款
    			tempRule=" and  Substr(Rule3,1,5) in ('"&RuleData&"') "
			elseif RuleType=4 then 
                '號
    			tempRule=" and  Substr(Rule3,1,7) in ('"&RuleData&"') "
			elseif RuleType=5 then 
                '全
    			tempRule=" and  Rule2 in ('"&RuleData&"') "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and  (Substr(Rule3,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) "
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and (Substr(Rule3,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) "
            end if   
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'" & tmpSql

			strSQL= strSQL& ")" 
			'response.write strSQL
			set rs2=conn.execute(strSQL)
			GetValue2=CDbl(rs2("Total"))
            rs2.close

            set rs2=nothing
end function 

function GetTotalValue(CarType,TypeID)

           ' BillTypeID TypeID 1攔停 2逕舉

           if CarType=1 then 
            '汽車
             TypeStr=" and CarSimpleID in ('2','1')  and (CarAddID not in ('9','5') or CarAddID is null)"
           elseif CarType=2 then 
            '輕機+一般重機
             TypeStr=" and CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
           elseif CarType=3 then 
            '大型重機
             TypeStr=" and (CarSimpleID='3' and CarNO Like 'FA%')"
           elseif CarType=4 then 
            '550CC重機
             TypeStr=" and CarAddID='9'"
           elseif CarType=5 then 
            '動力
             TypeStr=" and CarAddID='5'"
           end if
             

            strSQL= "select sum(total) as total from (" 
  		 	strSQL= strSQL&" select count(billno) as Total from BillBase  where Substr(Rule1,1,2) between '12' and '62' and BillTypeID='" & TypeID & "' " & TypeStr 
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql
			strSQL= strSQL& "union all "
  		 	strSQL= strSQL& " select count(billno) as Total from BillBase  where Substr(Rule2,1,2) between '12' and '62' and BillTypeID='" & TypeID & "' " & TypeStr 
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql		
			strSQL= strSQL& "union all "
  		 	strSQL= strSQL& " select count(billno) as Total from BillBase  where Substr(Rule3,1,2) between '12' and '62' and BillTypeID='" & TypeID & "' " & TypeStr 
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql				
			strSQL= strSQL& ")" 
			set rs2=conn.execute(strSQL)
			'response.write strsql
			GetTotalValue=CDbl(rs2("Total"))
            rs2.close
            set rs2=nothing
end function     

'慢車行人卻建1~68條的
function GetTotalValue2

            strSQL= "select sum(total) as total from (" 
  		 	strSQL= strSQL&" select count(billno) as Total from passerbase  where Substr(Rule1,1,2) between '12' and '62' "
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql
			strSQL= strSQL& "union all "
  		 	strSQL= strSQL& " select count(billno) as Total from passerbase  where Substr(Rule2,1,2) between '12' and '62' "
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql			
			strSQL= strSQL& "union all "
  		 	strSQL= strSQL& " select count(billno) as Total from passerbase  where Substr(Rule3,1,2) between '12' and '62' "
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and RecordStateID='0'"  & tmpSql			

			strSQL= strSQL& ")" 
			set rs2=conn.execute(strSQL)
			'response.write strsql
			GetTotalValue2=CDbl(rs2("Total"))
            rs2.close
            set rs2=nothing
end Function

function GetValue3(CarType,TypeID,RuleData,RuleType)
                          '   車種        攔逕     法條       法條式
           ' BillTypeID TypeID 1攔停 2逕舉
   	 lawList = RuleData
   	 lawList = Split(lawList,",")   	  
	 RuleData=""
   	 For i = 0 To UBound(lawList)
		tmpAry = lawList(i)
		tmpAry = Split(tmpAry,"$")
		if RuleData<>"" then RuleData=trim(RuleData)&"','"
		RuleData=trim(RuleData)&tmpAry(0)
   	 Next

           if CarType=1 then 
            '汽車
             TypeStr=" and CarSimpleID in ('2','1')  and (CarAddID not in ('9','5') or CarAddID is null)"
           elseif CarType=2 then 
            '輕機+一般重機
             TypeStr=" and CarSimpleID in ('4','3')  and (CarAddID not in ('9','5')  or CarAddID is null)"
           elseif CarType=3 then 
            '大型重機
             TypeStr=" and (CarSimpleID='3' and CarNO Like 'FA%')"
           elseif CarType=4 then 
            '550CC重機
             TypeStr=" and CarAddID='9'"
           elseif CarType=5 then 
            '動力
             TypeStr=" and CarAddID='5'"
           end if


            strSQL= "select sum(total) as total from (" 

			strSQL=strSQL&"Select count(billno) as Total from BillBase a,billfastenerdetail b"
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			
			if RuleType=1 then 
			    '條
    			tempRule=" and (Substr(Rule1,1,2) in ('"&RuleData&"') ) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and (Substr(Rule1,1,3) in ('"&RuleData&"') ) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule1,1,5) in ('"&RuleData&"') ) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and (Substr(Rule1,1,7) in ('"&RuleData&"') ) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule1 in ('"&RuleData&"') ) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ((Substr(Rule1,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule1,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule1,8,1) in ('"&right(RuleData,1)&"'))  )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and a.sn=b.billsn and a.carno=b.carno and b.fastenertypeid='6' and RecordStateID='0'" & tmpSql
			strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBase a,billfastenerdetail b"
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			if RuleType=1 then 
			    '條
    			tempRule=" and ( Substr(Rule2,1,2) in ('"&RuleData&"')) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and ( Substr(Rule2,1,3) in ('"&RuleData&"')) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule2,1,5) in ('"&RuleData&"')) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and ( Substr(Rule2,1,7) in ('"&RuleData&"')) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule2 in ('"&RuleData&"')) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ( (Substr(Rule2,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule2,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule2,8,1) in ('"&right(RuleData,1)&"')) )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and a.sn=b.billsn and a.carno=b.carno and b.fastenertypeid='6' and RecordStateID='0'" & tmpSql

			strSQL= strSQL& "union all "
			strSQL=strSQL&"Select count(billno) as Total from BillBase a,billfastenerdetail b"
			strSQL=strSQL & " where BillTypeID='" & TypeID & "' " & TypeStr 
			if RuleType=1 then 
			    '條
    			tempRule=" and ( Substr(Rule3,1,2) in ('"&RuleData&"')) "
			elseif RuleType=2 then 
			    '項
    			tempRule=" and ( Substr(Rule3,1,3) in ('"&RuleData&"')) "
			elseif RuleType=3 then 
                '款
    			tempRule=" and (Substr(Rule3,1,5) in ('"&RuleData&"')) "
			elseif RuleType=4 then 
                '號
    			tempRule=" and ( Substr(Rule3,1,7) in ('"&RuleData&"')) "
			elseif RuleType=5 then 
                '全
    			tempRule=" and (Rule3 in ('"&RuleData&"')) "
			elseif RuleType=6 then 
    			'之   ex:29-1
    			tempRule=" and ( (Substr(Rule3,1,2) in ('"&left(RuleData,2)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) )"
			elseif RuleType=7 then 
    			'項之  ex:293-1
    			tempRule=" and ((Substr(Rule3,1,3) in ('"&left(RuleData,3)&"') and Substr(Rule3,8,1) in ('"&right(RuleData,1)&"')) )"
            end if              
            
			strSQL=strSQL & tempRule
			strSQL= strSQL & " and " & DateType&" between TO_DATE('"&date1&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date2&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
			strSQL= strSQL& " and a.sn=b.billsn and a.carno=b.carno and b.fastenertypeid='6' and RecordStateID='0'" & tmpSql
			strSQL= strSQL& ")" 
'response.write strsql

			set rs2=conn.execute(strSQL)
			GetValue3=CDbl(rs2("Total"))
            rs2.close

            set rs2=nothing

end function  

%> 
<head>
<meta http-equiv=Content-Type content="text/html; charset=big5">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">
<link rel=Edit-Time-Data href="新修訂舉發報表-1.files/editdata.mso">
<link rel=OLE-Object-Data href="新修訂舉發報表-1.files/oledata.mso">

<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:0in 0in 0in 0in;
	mso-header-margin:0in;
	mso-footer-margin:0in;
	mso-page-orientation:landscape;
	mso-horizontal-page-align:center;
	mso-vertical-page-align:center;}
.font5
	{color:black;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;}
.font10
	{color:black;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
.font15
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"華康標楷體W6\(P\)", serif;
	mso-font-charset:136;}
.font17
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;}
tr
	{mso-height-source:auto;
	mso-ruby-visibility:none;}
col
	{mso-width-source:auto;
	mso-ruby-visibility:none;}
br
	{mso-data-placement:same-cell;}
.style0
	{mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	border:none;
	mso-protection:locked visible;
	mso-style-name:一般;
	mso-style-id:0;}
td
	{mso-style-parent:style0;
	padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xl24
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl25
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-protection:unlocked visible;}
.xl26
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl27
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl28
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center-across;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl29
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center-across;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl30
	{mso-style-parent:style0;
	color:black;
	font-size:16.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-protection:unlocked visible;}
.xl31
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl32
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl33
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl34
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl35
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl36
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl37
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	mso-protection:unlocked visible;}
.xl38
	{mso-style-parent:style0;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl39
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl40
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl41
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl42
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl43
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl44
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl45
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl46
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl47
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl48
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl49
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl50
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl51
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl52
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl53
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl54
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl55
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl56
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl57
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl58
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl59
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl60
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	mso-protection:unlocked visible;}
.xl61
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-protection:unlocked visible;}
.xl62
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl63
	{mso-style-parent:style0;
	color:black;
	font-size:20.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-protection:unlocked visible;}
.xl64
	{mso-style-parent:style0;
	color:black;
	font-size:8.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl65
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"m\0022月\0022d\0022日\0022";
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl66
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	mso-protection:unlocked visible;}
.xl67
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl68
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl69
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl70
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl71
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl72
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:細明體, monospace;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl73
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl74
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:left;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl75
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl76
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl77
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl78
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl79
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl80
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl81
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:left;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl82
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl83
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl84
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl85
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl86
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl87
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl88
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	mso-protection:unlocked visible;}
.xl89
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl90
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;
	white-space:normal;}
.xl91
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:"Times New Roman", serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl92
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl93
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl94
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl95
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl96
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl97
	{mso-style-parent:style0;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl98
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl99
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl100
	{mso-style-parent:style0;
	color:black;
	font-size:16.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl101
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	border:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto none;}
.xl102
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl103
	{mso-style-parent:style0;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl104
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl105
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border:.5pt solid windowtext;}
.xl106
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;}
.xl107
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl108
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl109
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl110
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;}
.xl111
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;}
.xl112
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;}
.xl113
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;}
.xl114
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;}
.xl115
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center-across;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl116
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl117
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl118
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl119
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:justify;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;
	white-space:normal;}
.xl120
	{mso-style-parent:style0;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl121
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl122
	{mso-style-parent:style0;
	color:black;
	font-size:9.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl123
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl124
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-protection:unlocked visible;}
.xl125
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:right;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;}
.xl126
	{mso-style-parent:style0;
	color:black;
	font-size:16.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center-across;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl127
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center-across;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl128
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl129
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl130
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl131
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl132
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl133
	{mso-style-parent:style0;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;}
.xl134
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-protection:unlocked visible;
	white-space:normal;}
.xl135
	{mso-style-parent:style0;
	color:black;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	vertical-align:justify;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl136
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	border-top:.5pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl137
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl138
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;
	white-space:normal;}
.xl139
	{mso-style-parent:style0;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	white-space:normal;}
.xl140
	{mso-style-parent:style0;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	white-space:normal;}
.xl141
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl142
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-protection:unlocked visible;}
.xl143
	{mso-style-parent:style0;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;}
.xl144
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-protection:unlocked visible;}
.xl145
	{mso-style-parent:style0;
	color:black;
	font-size:10.0pt;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	text-align:center;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;}
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
	mso-char-type:none;
	display:none;}
-->
</style>

</head>

<body link=blue vlink=purple class=xl61>

<table x:str border=0 cellpadding=0 cellspacing=0 width=2261 style='border-collapse:
 collapse;table-layout:fixed;width:1696pt'>
 <col class=xl25 width=85 style='mso-width-source:userset;mso-width-alt:2720;
 width:64pt'>
 <col class=xl25 width=68 span=29 style='mso-width-source:userset;mso-width-alt:
 2176;width:51pt'>
 <col class=xl25 width=68 style='mso-width-source:userset;mso-width-alt:2176;
 width:51pt'>
 <col class=xl61 width=68 span=225 style='mso-width-source:userset;mso-width-alt:
 2176;width:51pt'>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 width=85 style='height:15.0pt;width:64pt'>公 開 類</td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl25 width=68 style='width:51pt'></td>
  <td class=xl34 width=68 style='width:51pt'>　</td>
    <%
    sql = "select UnitName from UnitInfo where UnitID= '" & Session("Unit_ID") & "'"
    Set RSSystem = Conn.Execute(sql)
    if Not RSSystem.Eof Then
    	printUnit = RSSystem("UnitName")
    End If

If unit="y" And InStr(Sys_UnitID,"0460")=0 And InStr(Sys_UnitID,"0406")=0 Then
    sql = "select UnitName from UnitInfo where UnitID= '" & UnitID_Title & "'"
    Set RSSystem = Conn.Execute(sql)
    if Not RSSystem.Eof Then
    	TitleUnit = RSSystem("UnitName")
    End If
else
     TitleUnit=""
end if   


	unitList = Split(UnitID_q,"~")
SetUnit=""
		For i = 0 To UBound(unitList)
    sql = "select UnitName from UnitInfo where UnitID ='" & unitList(i) & "'"
    Set RSSystem = Conn.Execute(sql)

	While Not RSSystem.Eof
	  if trim(SetUnit)<>"" then 
	    SetUnit = SetUnit & "," & RSSystem("UnitName")
	  else
    	SetUnit = RSSystem("UnitName") 
      end if
       RSSystem.MoveNext
	Wend
			
		Next

	Set RSSystem=nothing
  %>
  <td class=xl24 width=68 style='border-left:none;width:51pt'>編報機關</td>
  <td class=xl26 width=68 style='border-left:none;width:51pt'><%=printUnit%></td>
  <td class=xl62 width=68 style='width:51pt'>　</td>
  <td class=xl62 width=68 style='width:51pt'>　</td>
  <td class=xl62 width=68 style='width:51pt'>　</td>
  <td class=xl27 width=68 style='width:51pt'>　</td>
  <td class=xl61 width=68 style='width:51pt'></td>
  <td class=xl61 width=68 style='width:51pt'></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl137 style='height:15.0pt;border-top:none'>月<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>報</td>
  <td class=xl25 colspan=3 style='mso-ignore:colspan'>每月終了後10日內編報</td>
  <td colspan=20 class=xl25 style='mso-ignore:colspan'></td>
  <td class=xl34>　</td>
  <td class=xl137 style='border-top:none;border-left:none'>表<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>號</td>
  <td colspan=5 class=xl141 style='border-right:1.0pt solid black;border-left:
  none'>1736 -01-01-2</td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl63 height=24 style='mso-height-source:userset;height:18.0pt'>
  <td height=24 colspan=31 class=xl126 align=center style='height:18.0pt;
  mso-ignore:colspan;border-right:1.0pt solid black'><%=TitleUnit%>        舉<span
  style='mso-spacerun:yes'>&nbsp; </span>發<span style='mso-spacerun:yes'>&nbsp;
  </span>違<span style='mso-spacerun:yes'>&nbsp; </span>反<span
  style='mso-spacerun:yes'>&nbsp; </span>道<span style='mso-spacerun:yes'>&nbsp;
  </span>路<span style='mso-spacerun:yes'>&nbsp; </span>交<span
  style='mso-spacerun:yes'>&nbsp; </span>通<span style='mso-spacerun:yes'>&nbsp;
  </span>管<span style='mso-spacerun:yes'>&nbsp; </span>理<span
  style='mso-spacerun:yes'>&nbsp; </span>事<span style='mso-spacerun:yes'>&nbsp;
  </span>件<span style='mso-spacerun:yes'>&nbsp; </span>成<span
  style='mso-spacerun:yes'>&nbsp; </span>果</td>
  <td colspan=2 class=xl63 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 colspan=26 class=xl127 align=center style='height:15.0pt;
  mso-ignore:colspan'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span><%=DateName%>中<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>華<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>民<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>國<span
  style='mso-spacerun:yes'>&nbsp;
  </span><%=CInt(Mid(date1,1,4))-1911%>年<%=datepart("m",date1)%>月<%=datepart("d",date1)%>日&nbsp;至&nbsp;<%=CInt(Mid(date2,1,4))-1911%>年<%=datepart("m",date2)%>月<%=datepart("d",date2)%>日</td>
  <td colspan=4 class=xl30 style='mso-ignore:colspan'></td>
  <td class=xl116>單位：件</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl128 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl31 style='border-top:none'>項</td>
  <td class=xl32 style='border-top:none;border-left:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>移</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>公</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>路</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl33>　</td>
  <td class=xl33>　</td>
  <td class=xl33>　</td>
  <td class=xl33>　</td>
  <td class=xl98 style='border-top:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>目</td>
  <td class=xl35 style='border-left:none'>舉</td>
  <td class=xl36>合</td>
  <td class=xl90 width=68 style='border-top:none;border-left:none;width:51pt'>未領</td>
  <td class=xl90 width=68 style='border-top:none;border-left:none;width:51pt'>拼</td>
  <td class=xl102 style='border-left:none'>報</td>
  <td class=xl102 style='border-left:none'>號牌</td>
  <td class=xl53 style='border-left:none'>損毀</td>
  <td class=xl53 style='border-left:none'>牌照</td>
  <td class=xl53 style='border-left:none'>各項</td>
  <td class=xl53 style='border-left:none'>設備</td>
  <td class=xl53 style='border-left:none'>重要設</td>
  <td class=xl53 style='border-left:none'>未</td>
  <td class=xl53 style='border-left:none'>駕</td>
  <td class=xl99 style='border-top:none;border-left:none'>駕</td>
  <td class=xl37>大型</td>
  <td class=xl53>大型</td>
  <td class=xl53 style='border-left:none'>大型</td>
  <td class=xl37>越</td>
  <td class=xl53>裝</td>
  <td class=xl53 style='border-left:none'>裝載</td>
  <td class=xl53 style='border-left:none'>裝載</td>
  <td class=xl90 width=68 style='border-top:none;border-left:none;width:51pt'>汽車裝</td>
  <td class=xl53 style='border-left:none'>載</td>
  <td class=xl53 style='border-left:none'>未</td>
  <td class=xl53 style='border-left:none'>未</td>
  <td class=xl90 width=68 style='border-top:none;border-left:none;width:51pt'>附載</td>
  <td class=xl53 style='border-left:none'>六歲</td>
  <td class=xl99 style='border-top:none;border-left:none'>機踏</td>
  <td class=xl34>未</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>與</td>
  <td class=xl35 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl102 style='border-left:none'>用或</td>
  <td class=xl102 style='border-left:none'>裝</td>
  <td class=xl102 style='border-left:none'>廢</td>
  <td class=xl102 style='border-left:none'>遺失</td>
  <td class=xl53 style='border-left:none'>或變</td>
  <td class=xl53 style='border-left:none'>遺失</td>
  <td class=xl53 style='border-left:none'>異動</td>
  <td class=xl53 style='border-left:none'>不全</td>
  <td class=xl53 style='border-left:none'>備變更</td>
  <td class=xl53 style='border-left:none'>領</td>
  <td class=xl53 style='border-left:none'>照</td>
  <td class=xl53 style='border-left:none'>照</td>
  <td class=xl37>車未</td>
  <td class=xl53>車駕</td>
  <td class=xl53 style='border-left:none'>車駕</td>
  <td class=xl37>級</td>
  <td class=xl53>載</td>
  <td class=xl53 style='border-left:none'>砂石</td>
  <td class=xl53 style='border-left:none'>超過</td>
  <td class=xl97 style='border-left:none'>載貨物</td>
  <td class=xl53 style='border-left:none'>運</td>
  <td class=xl53 style='border-left:none'>繫</td>
  <td class=xl53 style='border-left:none'>繫</td>
  <td class=xl53 style='border-left:none'>幼童</td>
  <td class=xl53 style='border-left:none'>以下</td>
  <td class=xl53 style='border-left:none'>車附</td>
  <td class=xl34>戴</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>適</td>
  <td class=xl35 style='border-left:none'>發</td>
  <td class=xl36>　</td>
  <td class=xl102 style='border-left:none'>未懸</td>
  <td class=xl102 style='border-left:none'>車</td>
  <td class=xl102 style='border-left:none'>汽</td>
  <td class=xl102 style='border-left:none'>經舉</td>
  <td class=xl53 style='border-left:none'>造牌</td>
  <td class=xl53 style='border-left:none'>或破</td>
  <td class=xl53 style='border-left:none'>不依</td>
  <td class=xl53 style='border-left:none'>或損</td>
  <td class=xl53 style='border-left:none'>或因事</td>
  <td class=xl53 style='border-left:none'>有</td>
  <td class=xl53 style='border-left:none'>不</td>
  <td class=xl53 style='border-left:none'>吊</td>
  <td class=xl37>領有</td>
  <td class=xl53>照不</td>
  <td class=xl53 style='border-left:none'>照吊</td>
  <td class=xl37>駕</td>
  <td class=xl53>不</td>
  <td class=xl53 style='border-left:none'>土方</td>
  <td class=xl53 style='border-left:none'>核定</td>
  <td class=xl97 style='border-left:none'>行經設</td>
  <td class=xl53 style='border-left:none'>客</td>
  <td class=xl53 style='border-left:none'>安</td>
  <td class=xl53 style='border-left:none'>安</td>
  <td class=xl53 style='border-left:none'>未依</td>
  <td class=xl53 style='border-left:none'>兒童</td>
  <td class=xl53 style='border-left:none'>載人</td>
  <td class=xl34>安</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>車</td>
  <td class=xl34>用</td>
  <td class=xl35 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl102 style='border-left:none'>掛牌</td>
  <td class=xl102 style='border-left:none'>輛</td>
  <td class=xl102 style='border-left:none'>車</td>
  <td class=xl102 style='border-left:none'>發後</td>
  <td class=xl53 style='border-left:none'>照致</td>
  <td class=xl53 style='border-left:none'>損不</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>壞不</td>
  <td class=xl53 style='border-left:none'>故損壞</td>
  <td class=xl53 style='border-left:none'>駕</td>
  <td class=xl53 style='border-left:none'>合</td>
  <td class=xl53 style='border-left:none'>扣</td>
  <td class=xl37>駕照</td>
  <td class=xl53>合規</td>
  <td class=xl53 style='border-left:none'>扣期</td>
  <td class=xl37>駛</td>
  <td class=xl53>合</td>
  <td class=xl53 style='border-left:none'>未依</td>
  <td class=xl53 style='border-left:none'>之總</td>
  <td class=xl97 style='border-left:none'>有地磅</td>
  <td class=xl53 style='border-left:none'>貨</td>
  <td class=xl53 style='border-left:none'>全</td>
  <td class=xl53 style='border-left:none'>全</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>單獨</td>
  <td class=xl53 style='border-left:none'>員或</td>
  <td class=xl34>全</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>輛</td>
  <td class=xl34>條</td>
  <td class=xl35 style='border-left:none'>總</td>
  <td class=xl36>　</td>
  <td class=xl102 style='border-left:none'>照等</td>
  <td class=xl102 style='border-left:none'>違</td>
  <td class=xl102 style='border-left:none'>仍</td>
  <td class=xl102 style='border-left:none'>不報</td>
  <td class=xl53 style='border-left:none'>不能</td>
  <td class=xl53 style='border-left:none'>報請</td>
  <td class=xl53 style='border-left:none'>申報</td>
  <td class=xl53 style='border-left:none'>予修</td>
  <td class=xl53 style='border-left:none'>未檢驗</td>
  <td class=xl53 style='border-left:none'>照</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>期</td>
  <td class=xl37>駕車</td>
  <td class=xl53>定者</td>
  <td class=xl53 style='border-left:none'>間駕</td>
  <td class=xl30></td>
  <td class=xl53>規</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>重量</td>
  <td class=xl97 style='border-left:none'>處所<font class="font17">1</font></td>
  <td class=xl53 style='border-left:none'>違</td>
  <td class=xl53 style='border-left:none'>帶</td>
  <td class=xl53 style='border-left:none'>帶</td>
  <td class=xl53 style='border-left:none'>安置</td>
  <td class=xl53 style='border-left:none'>留置</td>
  <td class=xl53 style='border-left:none'>物品</td>
  <td class=xl34>帽</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>與</td>
  <td class=xl34>例</td>
  <td class=xl35 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl102 style='border-left:none'>規</td>
  <td class=xl102 style='border-left:none'>行</td>
  <td class=xl102 style='border-left:none'>請補</td>
  <td class=xl53 style='border-left:none'>辨認</td>
  <td class=xl53 style='border-left:none'>補發</td>
  <td class=xl53 style='border-left:none'>登記</td>
  <td class=xl53 style='border-left:none'>復</td>
  <td class=xl53 style='border-left:none'>而行駛</td>
  <td class=xl53 style='border-left:none'>駕</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>間</td>
  <td class=xl133>　</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl30></td>
  <td class=xl53>定</td>
  <td class=xl53 style='border-left:none'>使用</td>
  <td class=xl53 style='border-left:none'>總聯</td>
  <td class=xl97 style='border-left:none'>公里內</td>
  <td class=xl53 style='border-left:none'>反</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>於安</td>
  <td class=xl53 style='border-left:none'>車內</td>
  <td class=xl53 style='border-left:none'>未依</td>
  <td class=xl34>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>舉</td>
  <td class=xl39>　</td>
  <td class=xl35 style='border-left:none'>件</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl102 style='border-left:none'>行</td>
  <td class=xl102 style='border-left:none'>駛</td>
  <td class=xl102 style='border-left:none'>發仍</td>
  <td class=xl53 style='border-left:none'>牌號</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl135 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>者</td>
  <td class=xl37>駕</td>
  <td class=xl103>　</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl25></td>
  <td class=xl53>　</td>
  <td class=xl53 style='border-left:none'>專用</td>
  <td class=xl53 style='border-left:none'>結重</td>
  <td class=xl97 style='border-left:none'>拒絕過</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl64 style='border-left:none'>(高速公路)</td>
  <td class=xl53 style='border-left:none'>全椅</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl34>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>發</td>
  <td class=xl39>　</td>
  <td class=xl35 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>12條第1項</td>
  <td class=xl102 style='border-left:none'>駛</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl102 style='border-left:none'>行駛</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl89 style='border-left:none'>　</td>
  <td class=xl135 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>21條第1項</td>
  <td class=xl37>車</td>
  <td class=xl53>　</td>
  <td class=xl53 style='border-left:none'>21條之1</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl30></td>
  <td class=xl53>　</td>
  <td class=xl53 style='border-left:none'>車輛</td>
  <td class=xl53 style='border-left:none'>量者</td>
  <td class=xl97 style='border-left:none'>磅</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl37></td>
  <td class=xl53>　</td>
  <td class=xl77 style='border-left:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>方</td>
  <td class=xl39>　</td>
  <td class=xl35 style='border-left:none'>數</td>
  <td class=xl36>　</td>
  <td class=xl37>第1.3.4.5</td>
  <td class=xl65>12條第1</td>
  <td class=xl65 style='border-left:none'>12條第1</td>
  <td class=xl65 style='border-left:none'>12條第1</td>
  <td class=xl53 style='border-left:none'>13條第1</td>
  <td class=xl53 style='border-left:none'>14條</td>
  <td class=xl53 style='border-left:none'>16條第1</td>
  <td class=xl53 style='border-left:none'>16條第1</td>
  <td class=xl53 style='border-left:none'>18條</td>
  <td class=xl53 style='border-left:none'>21條第</td>
  <td class=xl53 style='border-left:none'>第2.3.4</td>
  <td class=xl53 style='border-left:none'>21條第1</td>
  <td class=xl53 style='border-left:none'>21條之1第</td>
  <td class=xl53 style='border-left:none'>第1項第2.</td>
  <td class=xl53 style='border-left:none'>21條之1第</td>
  <td class=xl66>22<font class="font5">條</font></td>
  <td class=xl53>29條</td>
  <td class=xl53 style='border-left:none'>29條之1</td>
  <td class=xl53 style='border-left:none'>29條之2</td>
  <td class=xl53 style='border-left:none'>29條之2</td>
  <td class=xl53 style='border-left:none'>30條</td>
  <td class=xl53 style='border-left:none'>31條</td>
  <td class=xl53 style='border-left:none'>31條</td>
  <td class=xl53 style='border-left:none'>31條</td>
  <td class=xl37>31條</td>
  <td class=xl53>31條</td>
  <td class=xl77 style='border-left:none'>31條</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl130 style='height:15.0pt'>式</td>
  <td class=xl40>　</td>
  <td class=xl41 style='border-left:none'>　</td>
  <td class=xl42>計</td>
  <td class=xl54 style='border-left:none'>6.7.8款</td>
  <td class=xl54 style='border-left:none'>項第2款</td>
  <td class=xl54 style='border-left:none'>項第<font class="font10">9</font><font
  class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>項第<font class="font10">10</font><font
  class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>項第1款</td>
  <td class=xl57>第<font class="font10">1</font><font class="font5">款</font></td>
  <td class=xl54>項第<font class="font10">1</font><font class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>項第<font class="font10">2</font><font
  class="font5">款</font></td>
  <td class=xl57>第1項</td>
  <td class=xl54>1項第1款</td>
  <td class=xl67>6<font class="font5">.7.8.9款</font></td>
  <td class=xl54>項第<font class="font10">5</font><font class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>1項第1款</td>
  <td class=xl57>3.4.5.6款</td>
  <td class=xl54>1項第7款</td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">項</font></td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl54 style='border-left:none'>第1項</td>
  <td class=xl54 style='border-left:none'>第2項</td>
  <td class=xl54 style='border-left:none'>第3項</td>
  <td class=xl57>第4項</td>
  <td class=xl54>第5項</td>
  <td class=xl82 style='border-left:none'>第6項</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl144 style='border-right:1.0pt solid black;
  height:15.0pt'>合<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>計</td>
  <td class=xl107 align=right style='border-top:none;border-left:none' x:num x:fmla="=D16+U60">0</td>
  <td class=xl108 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(D17,D20,D23,D26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E17,E20,E23,E26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(F17,F20,F23,F26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(G17,G20,G23,G26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(H17,H20,H23,H26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(I17,I20,I23,I26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(J17,J20,J23,J26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(K17,K20,K23,K26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(L17,L20,L23,L26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(M17,M20,M23,M26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(N17,N20,N23,N26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(O17,O20,O23,O26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(P17,P20,P23,P26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Q17,Q20,Q23,Q26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(R17,R20,R23,R26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(S17,S20,S23,S26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(T17,T20,T23,T26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(U17,U20,U23,U26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(V17,V20,V23,V26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(W17,W20,W23,W26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(X17,X20,X23,X26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Y17,Y20,Y23,Y26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Z17,Z20,Z23,Z26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AA17,AA20,AA23,AA26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AB17,AB20,AB23,AB26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AC17,AC20,AC23,AC26)">0</td>
  <td class=xl104 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AD17,AD20,AD23,AD26)">0</td>
  <td class=xl109 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AE17,AE20,AE23,AE26)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:num x:fmla="=D17+U61">　</td>
  <td class=xl110 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(D18:D19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E18:E19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(F18:F19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(G18:G19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(H18:H19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(I18:I19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(J18:J19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(K18:K19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(L18:L19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(M18:M19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(N18:N19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(O18:O19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(P18:P19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Q18:Q19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(R18:R19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(S18:S19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(T18:T19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(U18:U19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(V18:V19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(W18:W19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(X18:X19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Y18:Y19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Z18:Z19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AA18:AA19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AB18:AB19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AC18:AC19)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AD18:AD19)">0</td>
  <td class=xl111 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AE18:AE19)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>汽車</td>
  <td class=xl43 style='border-top:none;border-left:none'>逕舉</td>
  <td class=xl136 align=right style='border-top:none;border-left:none'  x:fmla="=D18+U62">　</td>
  <td class=xl96 style='border-top:none;border-left:none' x:num x:fmla="=SUM(E18:AE18,C40:AE40,C62:T62,V62:AA62)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U18%>-SUM(V18:X18)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD18%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE18%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>攔停</td>
  <td class=xl136 align=right style='border-top:none;border-left:none'   x:fmla="=D19+U63">　</td>
  <td class=xl96 style='border-top:none;border-left:none' x:num x:fmla="=SUM(E19:AE19,C41:AE41,C63:T63,V63:AA63)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U19%>-SUM(V19:X19)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD19%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE19%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>550cc以上</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:num x:fmla="=D20+U64">　</td>
  <td class=xl110 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E20:AE20,C42:AE42,C64:T64,V64:AA64)">　</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E21:E22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(F21:F22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(G21:G22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(H21:H22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(I21:I22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(J21:J22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(K21:K22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(L21:L22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(M21:M22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(N21:N22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(O21:O22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(P21:P22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Q21:Q22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(R21:R22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(S21:S22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(T21:T22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(U21:U22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(V21:V22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(W21:W22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(X21:X22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Y21:Y22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Z21:Z22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AA21:AA22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AB21:AB22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AC21:AC22)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AD21:AD22)">0</td>
  <td class=xl111 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AE21:AE22)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>大型重</td>
  <td class=xl68 style='border-top:none'>逕舉</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:fmla="=D21+U65">　</td>
  <td class=xl96 style='border-top:none;border-left:none' x:num x:fmla="=SUM(E21:AE21,C43:AE43,C65:T65,V65:AA65)">　</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U21%>-SUM(V21:X21)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD21%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE21%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>型機車</td>
  <td class=xl68 style='border-top:none'>攔停</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:fmla="=D22+U66">　</td>
  <td class=xl96 style='border-top:none;border-left:none'  x:num x:fmla="=SUM(E22:AE22,C44:AE44,C66:T66,V66:AA66)">　</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U22%>-SUM(V21:X21)">　</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD22%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE22%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td rowspan=3 height=60 class=xl138 width=85 style='border-bottom:.5pt solid black;
  height:45.0pt;border-top:none;width:64pt'>未滿<br>
    550cc<br>
    機車</td>
  <td class=xl68 style='border-top:none'>小計</td>
  <td class=xl136 align=right style='border-top:none;border-left:none'  x:fmla="=D23+U67">　</td>
  <td class=xl110 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(D24:D25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E24:E25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(F24:F25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(G24:G25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(H24:H25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(I24:I25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(J24:J25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(K24:K25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(L24:L25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(M24:M25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(N24:N25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(O24:O25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(P24:P25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Q24:Q25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(R24:R25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(S24:S25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(T24:T25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(U24:U25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(V24:V25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(W24:W25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(X24:X25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Y24:Y25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(Z24:Z25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AA24:AA25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AB24:AB25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AC24:AC25)">0</td>
  <td class=xl105 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AD24:AD25)">0</td>
  <td class=xl111 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(AE24:AE25)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>逕舉</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:fmla="=D24+U68">　</td>
  <td class=xl96 style='border-top:none;border-left:none'  x:num x:fmla="=SUM(E24:AE24,C46:AE46,C68:T68,V68:AA68)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U24%>-SUM(V24:X24)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD24%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE24%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>攔停</td>
  <td class=xl136 align=right style='border-top:none;border-left:none' x:fmla="=D25+U69">　</td>
  <td class=xl96 style='border-top:none;border-left:none'  x:num x:fmla="=SUM(E25:AE25,C47:AE47,C69:T69,V69:AA69)">　</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=T25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="<%=U25%>-SUM(V24:X24)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=V25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=W25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=X25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Y25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Z25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AA25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AB25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AC25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AD25%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=AE25%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl95 style='height:15.0pt'>動力機械</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl106 align=right style='border-top:none;border-left:none' x:num  x:fmla="=D26+U70"></td>
  <td class=xl112 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(E26:AE26,C48:AE48,C70:T70,V70:AA70)">　</td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=E26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=F26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=G26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=H26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=I26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=J26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=K26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=L26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=M26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=N26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=O26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=P26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=Q26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=R26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=S26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=T26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num x:fmla="<%=U26%>-SUM(V24:X24)">　</td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=V26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=W26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=X26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=Y26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=Z26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=AA26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=AB26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=AC26%></td>
  <td class=xl113 align=right style='border-top:none;border-left:none' x:num><%=AD26%></td>
  <td class=xl114 align=right style='border-top:none;border-left:none' x:num><%=AE26%></td>
    <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl128 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl31>項</td>
  <td class=xl33 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>監</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl33 style='border-top:none'>理</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl49 style='border-top:none'>　</td>
  <td class=xl117 style='border-top:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>目</td>
  <td class=xl102>行駛</td>
  <td class=xl102 style='border-left:none'>動力</td>
  <td class=xl102 style='border-left:none'>非屬汽</td>
  <td class=xl102 style='border-left:none'>違反高<br>
    <br>
    <br>
    </td>
  <td class=xl53 style='border-left:none'>酒</td>
  <td class=xl53 style='border-left:none'>營業大</td>
  <td class=xl53 style='border-left:none'>酒後</td>
  <td class=xl53 style='border-left:none'>拒絕</td>
  <td class=xl53 style='border-left:none'>未</td>
  <td class=xl53 style='border-left:none'>不依</td>
  <td class=xl53 style='border-left:none'>執業</td>
  <td class=xl53 style='border-left:none'>違</td>
  <td class=xl53 style='border-left:none'>拒</td>
  <td class=xl53 style='border-left:none'>行車速</td>
  <td class=xl53 style='border-left:none'>蛇行</td>
  <td class=xl53 style='border-left:none'>行車速<br>
    <br>
    <font class="font5"><br>
    </font></td>
  <td class=xl53 style='border-left:none'>拆除</td>
  <td class=xl53 style='border-left:none'>二輛</td>
  <td class=xl53 style='border-left:none'>汽車所</td>
  <td class=xl53 style='border-left:none'>未</td>
  <td class=xl53 style='border-left:none'>行經</td>
  <td class=xl53 style='border-left:none'>不</td>
  <td class=xl53 style='border-left:none'>在多</td>
  <td class=xl53 style='border-left:none'>爭</td>
  <td class=xl53 style='border-left:none'>不</td>
  <td class=xl53 style='border-left:none'>不</td>
  <td class=xl53 style='border-left:none'>在多</td>
  <td class=xl53 style='border-left:none'>設有劃</td>
  <td class=xl76 style='border-top:none;border-left:none'>直行</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>與</td>
  <td class=xl102>道路</td>
  <td class=xl102 style='border-left:none'>機械</td>
  <td class=xl102 style='border-left:none'>車之載</td>
  <td class=xl102 style='border-left:none'>快速公</td>
  <td class=xl53 style='border-left:none'>後</td>
  <td class=xl53 style='border-left:none'>客車酒</td>
  <td class=xl53 style='border-left:none'>駕車</td>
  <td class=xl53 style='border-left:none'>接受</td>
  <td class=xl53 style='border-left:none'>辦</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>登記</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>載</td>
  <td class=xl53 style='border-left:none'>度超速</td>
  <td class=xl53 style='border-left:none'>或危</td>
  <td class=xl53 style='border-left:none'>度超速</td>
  <td class=xl53 style='border-left:none'>消音</td>
  <td class=xl53 style='border-left:none'>以上</td>
  <td class=xl53 style='border-left:none'>有人提</td>
  <td class=xl53 style='border-left:none'>依</td>
  <td class=xl53 style='border-left:none'>行人</td>
  <td class=xl53 style='border-left:none'>按</td>
  <td class=xl53 style='border-left:none'>車道</td>
  <td class=xl53 style='border-left:none'>道</td>
  <td class=xl53 style='border-left:none'>依</td>
  <td class=xl53 style='border-left:none'>依</td>
  <td class=xl53 style='border-left:none'>車道</td>
  <td class=xl53 style='border-left:none'>分島在</td>
  <td class=xl77 style='border-left:none'>車佔</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>適</td>
  <td class=xl102>時使</td>
  <td class=xl102 style='border-left:none'>未請</td>
  <td class=xl102 style='border-left:none'>具動力</td>
  <td class=xl102 style='border-left:none'>路管制</td>
  <td class=xl53 style='border-left:none'>駕</td>
  <td class=xl53 style='border-left:none'>後或吸</td>
  <td class=xl53 style='border-left:none'>吊扣</td>
  <td class=xl53 style='border-left:none'>酒精</td>
  <td class=xl53 style='border-left:none'>理</td>
  <td class=xl53 style='border-left:none'>辦理</td>
  <td class=xl53 style='border-left:none'>證未</td>
  <td class=xl53 style='border-left:none'>攬</td>
  <td class=xl53 style='border-left:none'>短</td>
  <td class=xl53 style='border-left:none'>60公里</td>
  <td class=xl53 style='border-left:none'>險駕</td>
  <td class=xl53 style='border-left:none'>60公里</td>
  <td class=xl53 style='border-left:none'>器或</td>
  <td class=xl53 style='border-left:none'>競駛</td>
  <td class=xl53 style='border-left:none'>供汽車</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>穿越</td>
  <td class=xl53 style='border-left:none'>遵</td>
  <td class=xl53 style='border-left:none'>不依</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>轉彎</td>
  <td class=xl53 style='border-left:none'>慢車道</td>
  <td class=xl77 style='border-left:none'>用轉</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>車</td>
  <td class=xl34>用</td>
  <td class=xl102>用手</td>
  <td class=xl102 style='border-left:none'>領臨</td>
  <td class=xl102 style='border-left:none'>休閒器</td>
  <td class=xl102 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>食管制</td>
  <td class=xl53 style='border-left:none'>駕照</td>
  <td class=xl53 style='border-left:none'>或管</td>
  <td class=xl53 style='border-left:none'>執</td>
  <td class=xl53 style='border-left:none'>異動</td>
  <td class=xl53 style='border-left:none'>依規</td>
  <td class=xl53 style='border-left:none'>客</td>
  <td class=xl53 style='border-left:none'>程</td>
  <td class=xl53 style='border-left:none'>以下</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>以上</td>
  <td class=xl53 style='border-left:none'>以其</td>
  <td class=xl53 style='border-left:none'>競技</td>
  <td class=xl53 style='border-left:none'>駕駛人</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>道不</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>駛</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>不依</td>
  <td class=xl53 style='border-left:none'>右轉彎</td>
  <td class=xl77 style='border-left:none'>彎專</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>輛</td>
  <td class=xl34>條</td>
  <td class=xl102>持式</td>
  <td class=xl102 style='border-left:none'>時通</td>
  <td class=xl102 style='border-left:none'>材違規</td>
  <td class=xl102 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>藥品後</td>
  <td class=xl53 style='border-left:none'>期間</td>
  <td class=xl53 style='border-left:none'>制藥</td>
  <td class=xl53 style='border-left:none'>業</td>
  <td class=xl53 style='border-left:none'>申請</td>
  <td class=xl53 style='border-left:none'>定安</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>繞</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>他方</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>危險駕</td>
  <td class=xl53 style='border-left:none'>減</td>
  <td class=xl53 style='border-left:none'>暫停</td>
  <td class=xl53 style='border-left:none'>方</td>
  <td class=xl53 style='border-left:none'>駕車</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>超</td>
  <td class=xl53 style='border-left:none'>轉</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>或在快</td>
  <td class=xl77 style='border-left:none'>用車</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>與</td>
  <td class=xl34>例</td>
  <td class=xl102>行動</td>
  <td class=xl102 style='border-left:none'>行證</td>
  <td class=xl102 style='border-left:none'>行駛</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>駕車</td>
  <td class=xl53 style='border-left:none'>仍駕</td>
  <td class=xl53 style='border-left:none'>品之</td>
  <td class=xl53 style='border-left:none'>登</td>
  <td class=xl53 style='border-left:none'>或年</td>
  <td class=xl53 style='border-left:none'>置等</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>道</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>式製</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>速</td>
  <td class=xl53 style='border-left:none'>讓行</td>
  <td class=xl53 style='border-left:none'>向</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>彎</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車道左</td>
  <td class=xl77 style='border-left:none'>道</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>舉</td>
  <td class=xl39>　</td>
  <td class=xl102>電話</td>
  <td class=xl37></td>
  <td class=xl134 width=68 style='width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車者</td>
  <td class=xl53 style='border-left:none'>檢測</td>
  <td class=xl53 style='border-left:none'>記</td>
  <td class=xl53 style='border-left:none'>度查</td>
  <td class=xl37></td>
  <td class=xl53>　</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>造噪</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>慢</td>
  <td class=xl53 style='border-left:none'>人先行</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>轉彎</td>
  <td class=xl77 style='border-left:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>發</td>
  <td class=xl39>　</td>
  <td class=xl94 style='border-left:none'>　</td>
  <td class=xl37></td>
  <td class=xl134 width=68 style='width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl103 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>驗者</td>
  <td class=xl37></td>
  <td class=xl53>　</td>
  <td class=xl53 style='border-left:none'>駛</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>音</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl38 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl134 width=68 style='border-left:none;width:51pt'>　</td>
  <td class=xl53 style='border-left:none'>駛</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl69 style='border-left:none'>45<font class="font5">條</font></td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>48<font class="font5">條</font></td>
  <td class=xl53 style='border-left:none'>48<font class="font5">條</font></td>
  <td class=xl37>48<font class="font5">條</font></td>
  <td class=xl87>48<font class="font5">條</font></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>方</td>
  <td class=xl39>　</td>
  <td class=xl94 style='border-left:none'>31條之1</td>
  <td class=xl37>32條</td>
  <td class=xl53>32條之<font class="font10">1</font></td>
  <td class=xl53 style='border-left:none'>33條</td>
  <td class=xl53 style='border-left:none'>35條</td>
  <td class=xl69 style='border-left:none'>3<font class="font5">5條</font></td>
  <td class=xl53 style='border-left:none'>35條</td>
  <td class=xl53 style='border-left:none'>35條</td>
  <td class=xl53 style='border-left:none'>36條</td>
  <td class=xl53 style='border-left:none'>36條</td>
  <td class=xl37>36條</td>
  <td class=xl53>38條</td>
  <td class=xl53 style='border-left:none'>38條</td>
  <td class=xl53 style='border-left:none'>40條</td>
  <td class=xl52 style='border-left:none'>43<font class="font5">條第</font><font
  class="font10">1</font></td>
  <td class=xl52>43<font class="font5">條第</font><font class="font10">1</font></td>
  <td class=xl52>43<font class="font5">條第</font><font class="font10">1</font></td>
  <td class=xl53>43條</td>
  <td class=xl53 style='border-left:none'>43條</td>
  <td class=xl53 style='border-left:none'>44條</td>
  <td class=xl53 style='border-left:none'>44條</td>
  <td class=xl53 style='border-left:none'>45條</td>
  <td class=xl53 style='border-left:none'>45條</td>
  <td class=xl53 style='border-left:none'>第2.3款</td>
  <td class=xl53 style='border-left:none'>47條</td>
  <td class=xl53 style='border-left:none'>第<font class="font10">1</font><font
  class="font5">項第</font><font class="font10">1.2</font></td>
  <td class=xl53 style='border-left:none'>第<font class="font10">1</font><font
  class="font5">項</font></td>
  <td class=xl53 style='border-left:none'>第<font class="font10">1</font><font
  class="font5">項</font></td>
  <td class=xl77 style='border-left:none'>第<font class="font10">1</font><font
  class="font5">項</font></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl130 style='height:15.0pt'>式</td>
  <td class=xl40>　</td>
  <td class=xl95 style='border-left:none'>　</td>
  <td class=xl57>第<font class="font10">1</font><font class="font5">項</font></td>
  <td class=xl54>　</td>
  <td class=xl54 style='border-left:none'>　</td>
  <td class=xl54 style='border-left:none'>第1項</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">2</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第1項</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl57>第5項</td>
  <td class=xl54>第1項</td>
  <td class=xl54 style='border-left:none'>第2項</td>
  <td class=xl56>　</td>
  <td class=xl70>項第<font class="font10">1</font><font class="font5">款</font></td>
  <td class=xl70>項第<font class="font10">2</font><font class="font5">款</font></td>
  <td class=xl70>項第<font class="font10">3</font><font class="font5">款</font></td>
  <td class=xl54>第3項</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">1</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第2項</td>
  <td class=xl54 style='border-left:none'>第1款</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">5~15</font><font
  class="font5">款</font></td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl71 style='border-left:none'>3.6<font class="font5">款</font></td>
  <td class=xl72 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">款</font></td>
  <td class=xl72 style='border-left:none'>第<font class="font10">5</font><font
  class="font5">款</font></td>
  <td class=xl82 style='border-left:none'>第<font class="font10">7</font><font
  class="font5">款</font></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl144 style='border-right:1.0pt solid black;
  height:15.0pt'>合計</td>
  <td class=xl50 align=right x:num x:fmla="=SUM(C39,C42,C45,C48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(D39,D42,D45,D48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(E39,E42,E45,E48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(F39,F42,F45,F48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(G39,G42,G45,G48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(H39,H42,H45,H48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(I39,I42,I45,I48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(J39,J42,J45,J48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(K39,K42,K45,K48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(L39,L42,L45,L48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(M39,M42,M45,M48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(N39,N42,N45,N48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(O39,O42,O45,O48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(P39,P42,P45,P48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Q39,Q42,Q45,Q48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(R39,R42,R45,R48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(S39,S42,S45,S48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(T39,T42,T45,T48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(U39,U42,U45,U48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(V39,V42,V45,V48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(W39,W42,W45,W48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(X39,X42,X45,X48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Y39,Y42,Y45,Y48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Z39,Z42,Z45,Z48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AA39,AA42,AA45,AA48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AB39,AB42,AB45,AB48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AC39,AC42,AC45,AC48)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AD39,AD42,AD45,AD48)">0</td>
  <td class=xl118 align=right style='border-left:none' x:num x:fmla="=SUM(AE39,AE42,AE45,AE48)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl50 align=right x:num x:fmla="=SUM(C40:C41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(D40:D41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(E40:E41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(F40:F41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(G40:G41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(H40:H41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(I40:I41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(J40:J41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(K40:K41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(L40:L41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(M40:M41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(N40:N41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(O40:O41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(P40:P41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Q40:Q41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(R40:R41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(S40:S41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(T40:T41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(U40:U41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(V40:V41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(W40:W41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(X40:X41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Y40:Y41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Z40:Z41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AA40:AA41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AB40:AB41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AC40:AC41)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AD40:AD41)">0</td>
  <td class=xl118 align=right style='border-left:none' x:num  x:fmla="=SUM(AE40:AE41)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>汽車</td>
  <td class=xl43 style='border-top:none;border-left:none'>逕舉</td>
  <td class=xl50 x:num><%=C40%></td>
  <td class=xl50 style='border-left:none' x:num><%=D40%></td>
  <td class=xl50 style='border-left:none' x:num><%=E40%></td>
  <td class=xl50 style='border-left:none' x:num><%=F40%></td>
  <td class=xl50 style='border-left:none' x:num><%=G40%></td>
  <td class=xl50 style='border-left:none' x:num><%=H40%></td>
  <td class=xl50 style='border-left:none' x:num><%=I40%></td>
  <td class=xl50 style='border-left:none' x:num><%=J40%></td>
  <td class=xl50 style='border-left:none' x:num><%=K40%></td>
  <td class=xl50 style='border-left:none' x:num><%=L40%></td>
  <td class=xl50 style='border-left:none' x:num><%=M40%></td>
  <td class=xl50 style='border-left:none' x:num><%=N40%></td>
  <td class=xl50 style='border-left:none' x:num><%=O40%></td>
  <td class=xl50 style='border-left:none' x:num><%=P40%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q40%></td>
  <td class=xl50 style='border-left:none' x:num><%=R40%></td>
  <td class=xl50 style='border-left:none' x:num><%=S40%></td>
  <td class=xl50 style='border-left:none' x:num><%=T40%></td>
  <td class=xl50 style='border-left:none' x:num><%=U40%></td>
  <td class=xl50 style='border-left:none' x:num><%=V40%></td>
  <td class=xl50 style='border-left:none' x:num><%=W40%></td>
  <td class=xl50 style='border-left:none' x:num><%=X40%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y40%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z40%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA40%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB40%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC40%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD40%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE40%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>攔停</td>
  <td class=xl50 x:num><%=C41%></td>
  <td class=xl50 style='border-left:none' x:num><%=D41%></td>
  <td class=xl50 style='border-left:none' x:num><%=E41%></td>
  <td class=xl50 style='border-left:none' x:num><%=F41%></td>
  <td class=xl50 style='border-left:none' x:num><%=G41%></td>
  <td class=xl50 style='border-left:none' x:num><%=H41%></td>
  <td class=xl50 style='border-left:none' x:num><%=I41%></td>
  <td class=xl50 style='border-left:none' x:num><%=J41%></td>
  <td class=xl50 style='border-left:none' x:num><%=K41%></td>
  <td class=xl50 style='border-left:none' x:num><%=L41%></td>
  <td class=xl50 style='border-left:none' x:num><%=M41%></td>
  <td class=xl50 style='border-left:none' x:num><%=N41%></td>
  <td class=xl50 style='border-left:none' x:num><%=O41%></td>
  <td class=xl50 style='border-left:none' x:num><%=P41%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q41%></td>
  <td class=xl50 style='border-left:none' x:num><%=R41%></td>
  <td class=xl50 style='border-left:none' x:num><%=S41%></td>
  <td class=xl50 style='border-left:none' x:num><%=T41%></td>
  <td class=xl50 style='border-left:none' x:num><%=U41%></td>
  <td class=xl50 style='border-left:none' x:num><%=V41%></td>
  <td class=xl50 style='border-left:none' x:num><%=W41%></td>
  <td class=xl50 style='border-left:none' x:num><%=X41%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y41%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z41%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA41%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB41%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC41%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD41%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE41%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>550cc以上</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl50 align=right x:num x:fmla="=SUM(C43:C44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(D43:D44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(E43:E44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(F43:F44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(G43:G44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(H43:H44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(I43:I44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(J43:J44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(K43:K44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(L43:L44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(M43:M44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(N43:N44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(O43:O44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(P43:P44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Q43:Q44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(R43:R44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(S43:S44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(T43:T44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(U43:U44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(V43:V44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(W43:W44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(X43:X44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Y43:Y44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(Z43:Z44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AA43:AA44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AB43:AB44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AC43:AC44)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num  x:fmla="=SUM(AD43:AD44)">0</td>
  <td class=xl118 align=right style='border-left:none' x:num  x:fmla="=SUM(AE43:AE44)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>大型重</td>
  <td class=xl43 style='border-top:none;border-left:none'>逕舉</td>
  <td class=xl50 x:num><%=C43%></td>
  <td class=xl50 style='border-left:none' x:num><%=D43%></td>
  <td class=xl50 style='border-left:none' x:num><%=E43%></td>
  <td class=xl50 style='border-left:none' x:num><%=F43%></td>
  <td class=xl50 style='border-left:none' x:num><%=G43%></td>
  <td class=xl50 style='border-left:none' x:num><%=H43%></td>
  <td class=xl50 style='border-left:none' x:num><%=I43%></td>
  <td class=xl50 style='border-left:none' x:num><%=J43%></td>
  <td class=xl50 style='border-left:none' x:num><%=K43%></td>
  <td class=xl50 style='border-left:none' x:num><%=L43%></td>
  <td class=xl50 style='border-left:none' x:num><%=M43%></td>
  <td class=xl50 style='border-left:none' x:num><%=N43%></td>
  <td class=xl50 style='border-left:none' x:num><%=O43%></td>
  <td class=xl50 style='border-left:none' x:num><%=P43%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q43%></td>
  <td class=xl50 style='border-left:none' x:num><%=R43%></td>
  <td class=xl50 style='border-left:none' x:num><%=S43%></td>
  <td class=xl50 style='border-left:none' x:num><%=T43%></td>
  <td class=xl50 style='border-left:none' x:num><%=U43%></td>
  <td class=xl50 style='border-left:none' x:num><%=V43%></td>
  <td class=xl50 style='border-left:none' x:num><%=W43%></td>
  <td class=xl50 style='border-left:none' x:num><%=X43%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y43%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z43%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA43%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB43%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC43%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD43%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE43%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>型機車</td>
  <td class=xl43 style='border-top:none;border-left:none'>攔停</td>
  <td class=xl50 x:num><%=C44%>
  <td class=xl50 style='border-left:none' x:num><%=D44%>
  <td class=xl50 style='border-left:none' x:num><%=E44%>
  <td class=xl50 style='border-left:none' x:num><%=F44%>
  <td class=xl50 style='border-left:none' x:num><%=G44%>
  <td class=xl50 style='border-left:none' x:num><%=H44%>
  <td class=xl50 style='border-left:none' x:num><%=I44%>
  <td class=xl50 style='border-left:none' x:num><%=J44%>
  <td class=xl50 style='border-left:none' x:num><%=K44%>
  <td class=xl50 style='border-left:none' x:num><%=L44%>
  <td class=xl50 style='border-left:none' x:num><%=M44%>
  <td class=xl50 style='border-left:none' x:num><%=N44%>
  <td class=xl50 style='border-left:none' x:num><%=O44%>
  <td class=xl50 style='border-left:none' x:num><%=P44%>
  <td class=xl50 style='border-left:none' x:num><%=Q44%>
  <td class=xl50 style='border-left:none' x:num><%=R44%>
  <td class=xl50 style='border-left:none' x:num><%=S44%>
  <td class=xl50 style='border-left:none' x:num><%=T44%>
  <td class=xl50 style='border-left:none' x:num><%=U44%>
  <td class=xl50 style='border-left:none' x:num><%=V44%>
  <td class=xl50 style='border-left:none' x:num><%=W44%>
  <td class=xl50 style='border-left:none' x:num><%=X44%>
  <td class=xl50 style='border-left:none' x:num><%=Y44%>
  <td class=xl50 style='border-left:none' x:num><%=Z44%>
  <td class=xl50 style='border-left:none' x:num><%=AA44%>
  <td class=xl50 style='border-left:none' x:num><%=AB44%>
  <td class=xl50 style='border-left:none' x:num><%=AC44%>
  <td class=xl50 style='border-left:none' x:num><%=AD44%>
  <td class=xl118 style='border-left:none' x:num><%=AE44%>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td rowspan=3 height=60 class=xl138 width=85 style='border-bottom:.5pt solid black;
  height:45.0pt;border-top:none;width:64pt'>未滿<br>
    550cc<br>
    機車</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl50 align=right x:num x:fmla="=SUM(C46:C47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(D46:D47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(E46:E47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(F46:F47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(G46:G47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(H46:H47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(I46:I47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(J46:J47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(K46:K47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(L46:L47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(M46:M47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(N46:N47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(O46:O47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(P46:P47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(Q46:Q47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(R46:R47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(S46:S47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(T46:T47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(U46:U47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(V46:V47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(W46:W47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(X46:X47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(Y46:Y47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(Z46:Z47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(AA46:AA47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(AB46:AB47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(AC46:AC47)">0</td>
  <td class=xl50 align=right style='border-left:none' x:num x:fmla="=SUM(AD46:AD47)">0</td>
  <td class=xl118 align=right style='border-left:none' x:num x:fmla="=SUM(AE46:AE47)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>逕舉</td>
  <td class=xl50 x:num><%=C46%></td>
  <td class=xl50 style='border-left:none' x:num><%=D46%></td>
  <td class=xl50 style='border-left:none' x:num><%=E46%></td>
  <td class=xl50 style='border-left:none' x:num><%=F46%></td>
  <td class=xl50 style='border-left:none' x:num><%=G46%></td>
  <td class=xl50 style='border-left:none' x:num><%=H46%></td>
  <td class=xl50 style='border-left:none' x:num><%=I46%></td>
  <td class=xl50 style='border-left:none' x:num><%=J46%></td>
  <td class=xl50 style='border-left:none' x:num><%=K46%></td>
  <td class=xl50 style='border-left:none' x:num><%=L46%></td>
  <td class=xl50 style='border-left:none' x:num><%=M46%></td>
  <td class=xl50 style='border-left:none' x:num><%=N46%></td>
  <td class=xl50 style='border-left:none' x:num><%=O46%></td>
  <td class=xl50 style='border-left:none' x:num><%=P46%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q%></td>
  <td class=xl50 style='border-left:none' x:num><%=R46%></td>
  <td class=xl50 style='border-left:none' x:num><%=S46%></td>
  <td class=xl50 style='border-left:none' x:num><%=T46%></td>
  <td class=xl50 style='border-left:none' x:num><%=U46%></td>
  <td class=xl50 style='border-left:none' x:num><%=V46%></td>
  <td class=xl50 style='border-left:none' x:num><%=W46%></td>
  <td class=xl50 style='border-left:none' x:num><%=X46%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y46%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z46%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA46%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB46%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC46%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD46%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE46%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>攔停</td>
  <td class=xl50 x:num><%=C47%></td>
  <td class=xl50 style='border-left:none' x:num><%=D47%></td>
  <td class=xl50 style='border-left:none' x:num><%=E47%></td>
  <td class=xl50 style='border-left:none' x:num><%=F47%></td>
  <td class=xl50 style='border-left:none' x:num><%=G47%></td>
  <td class=xl50 style='border-left:none' x:num><%=H47%></td>
  <td class=xl50 style='border-left:none' x:num><%=I47%></td>
  <td class=xl50 style='border-left:none' x:num><%=J47%></td>
  <td class=xl50 style='border-left:none' x:num><%=K47%></td>
  <td class=xl50 style='border-left:none' x:num><%=L47%></td>
  <td class=xl50 style='border-left:none' x:num><%=M47%></td>
  <td class=xl50 style='border-left:none' x:num><%=N47%></td>
  <td class=xl50 style='border-left:none' x:num><%=O47%></td>
  <td class=xl50 style='border-left:none' x:num><%=P47%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q47%></td>
  <td class=xl50 style='border-left:none' x:num><%=R47%></td>
  <td class=xl50 style='border-left:none' x:num><%=S47%></td>
  <td class=xl50 style='border-left:none' x:num><%=T47%></td>
  <td class=xl50 style='border-left:none' x:num><%=U47%></td>
  <td class=xl50 style='border-left:none' x:num><%=V47%></td>
  <td class=xl50 style='border-left:none' x:num><%=W47%></td>
  <td class=xl50 style='border-left:none' x:num><%=X47%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y47%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z47%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA47%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB47%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC47%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD47%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE47%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl95 style='height:15.0pt'>動力機械</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl50 x:num><%=C48%></td>
  <td class=xl50 style='border-left:none' x:num><%=D48%></td>
  <td class=xl50 style='border-left:none' x:num><%=E48%></td>
  <td class=xl50 style='border-left:none' x:num><%=F48%></td>
  <td class=xl50 style='border-left:none' x:num><%=G48%></td>
  <td class=xl50 style='border-left:none' x:num><%=H48%></td>
  <td class=xl50 style='border-left:none' x:num><%=I48%></td>
  <td class=xl50 style='border-left:none' x:num><%=J48%></td>
  <td class=xl50 style='border-left:none' x:num><%=K48%></td>
  <td class=xl50 style='border-left:none' x:num><%=L48%></td>
  <td class=xl50 style='border-left:none' x:num><%=M48%></td>
  <td class=xl50 style='border-left:none' x:num><%=N48%></td>
  <td class=xl50 style='border-left:none' x:num><%=O48%></td>
  <td class=xl50 style='border-left:none' x:num><%=P48%></td>
  <td class=xl50 style='border-left:none' x:num><%=Q48%></td>
  <td class=xl50 style='border-left:none' x:num><%=R48%></td>
  <td class=xl50 style='border-left:none' x:num><%=S48%></td>
  <td class=xl50 style='border-left:none' x:num><%=T48%></td>
  <td class=xl50 style='border-left:none' x:num><%=U48%></td>
  <td class=xl50 style='border-left:none' x:num><%=V48%></td>
  <td class=xl50 style='border-left:none' x:num><%=W48%></td>
  <td class=xl50 style='border-left:none' x:num><%=X48%></td>
  <td class=xl50 style='border-left:none' x:num><%=Y48%></td>
  <td class=xl50 style='border-left:none' x:num><%=Z48%></td>
  <td class=xl50 style='border-left:none' x:num><%=AA48%></td>
  <td class=xl50 style='border-left:none' x:num><%=AB48%></td>
  <td class=xl50 style='border-left:none' x:num><%=AC48%></td>
  <td class=xl50 style='border-left:none' x:num><%=AD48%></td>
  <td class=xl118 style='border-left:none' x:num><%=AE48%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl128 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl31>項</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl33>機</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl49>　</td>
  <td class=xl33>關</td>
  <td class=xl51>　</td>
  <td class=xl73>　</td>
  <td class=xl100 style='border-left:none'>　</td>
  <td class=xl51>警</td>
  <td class=xl51>　</td>
  <td class=xl74>察</td>
  <td class=xl74>機</td>
  <td class=xl74>關</td>
  <td class=xl98>　</td>
  <td class=xl33>其</td>
  <td class=xl33>　</td>
  <td class=xl75>　</td>
  <td class=xl31>他</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>目</td>
  <td class=xl93 style='border-top:none;border-left:none'>轉彎</td>
  <td class=xl36>闖</td>
  <td class=xl99 style='border-top:none;border-left:none'>闖</td>
  <td class=xl37>闖越</td>
  <td class=xl53>違</td>
  <td class=xl53 style='border-left:none'>違</td>
  <td class=xl53 style='border-left:none'>停車</td>
  <td class=xl53 style='border-left:none'>於身</td>
  <td class=xl99 style='border-top:none;border-left:none'>在道</td>
  <td class=xl37>路</td>
  <td class=xl53>不服</td>
  <td class=xl53 style='border-left:none'>其他</td>
  <td class=xl53 style='border-left:none'>違反</td>
  <td class=xl53 style='border-left:none'>肇事</td>
  <td class=xl53 style='border-left:none'>肇事</td>
  <td class=xl53 style='border-left:none'>肇事</td>
  <td class=xl53 style='border-left:none'>肇事</td>
  <td class=xl76 style='border-top:none;border-left:none'>其他</td>
  <td class=xl36>合</td>
  <td class=xl53 style='border-left:none'>慢</td>
  <td class=xl53 style='border-left:none'>行</td>
  <td class=xl53 style='border-left:none'>在道</td>
  <td class=xl99 style='border-top:none;border-left:none'>未經</td>
  <td class=xl37>在車道交</td>
  <td class=xl87>82<font class="font5">條第</font><font class="font10">1</font><font
  class="font5">項</font></td>
  <td class=xl53>強制</td>
  <td class=xl53 style='border-left:none'>違</td>
  <td class=xl99 style='border-left:none'>查報</td>
  <td class=xl119 width=68 style='border-left:none;width:51pt'>違規<br>
    車輛<br>
    移置<br>
    保管</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>與</td>
  <td class=xl94 style='border-left:none'>不暫</td>
  <td class=xl36>紅</td>
  <td class=xl53 style='border-left:none'>紅</td>
  <td class=xl37>平交</td>
  <td class=xl53>規</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>時間</td>
  <td class=xl53 style='border-left:none'>心障</td>
  <td class=xl53 style='border-left:none'>路上</td>
  <td class=xl37>口</td>
  <td class=xl53>或抗</td>
  <td class=xl53 style='border-left:none'>不遵</td>
  <td class=xl53 style='border-left:none'>道安</td>
  <td class=xl53 style='border-left:none'>無人</td>
  <td class=xl53 style='border-left:none'>無人</td>
  <td class=xl53 style='border-left:none'>致人</td>
  <td class=xl53 style='border-left:none'>致人</td>
  <td class=xl77 style='border-left:none'>汽機</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>人</td>
  <td class=xl53 style='border-left:none'>路上</td>
  <td class=xl53 style='border-left:none'>許可</td>
  <td class=xl37>通島散發</td>
  <td class=xl77>第<font class="font10">1.10</font><font class="font5">款</font></td>
  <td class=xl53>移由</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>佔用</td>
  <td class=xl120 style='border-left:none'>車輛</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>　</td>
  <td class=xl34>適</td>
  <td class=xl94 style='border-left:none'>停讓</td>
  <td class=xl36>燈</td>
  <td class=xl53 style='border-left:none'>燈</td>
  <td class=xl37>道或</td>
  <td class=xl53>臨</td>
  <td class=xl53 style='border-left:none'>停</td>
  <td class=xl53 style='border-left:none'>位置</td>
  <td class=xl53 style='border-left:none'>礙專</td>
  <td class=xl53 style='border-left:none'>停放</td>
  <td class=xl37>淨</td>
  <td class=xl53>拒交</td>
  <td class=xl53 style='border-left:none'>守標</td>
  <td class=xl53 style='border-left:none'>規則</td>
  <td class=xl53 style='border-left:none'>傷亡</td>
  <td class=xl53 style='border-left:none'>傷亡</td>
  <td class=xl53 style='border-left:none'>死傷</td>
  <td class=xl53 style='border-left:none'>傷亡</td>
  <td class=xl77 style='border-left:none'>車違</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>違</td>
  <td class=xl53 style='border-left:none'>違 </td>
  <td class=xl53 style='border-left:none'>堆積</td>
  <td class=xl53 style='border-left:none'>在道</td>
  <td class=xl37>廣告物等</td>
  <td class=xl87>83<font class="font5">條第</font><font class="font10">1</font><font
  class="font5">項</font></td>
  <td class=xl53>醫療</td>
  <td class=xl53 style='border-left:none'>停</td>
  <td class=xl53 style='border-left:none'>道路</td>
  <td class=xl120 style='border-left:none'>移置</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>車</td>
  <td class=xl34>用</td>
  <td class=xl94 style='border-left:none'>行人</td>
  <td class=xl36>直</td>
  <td class=xl53 style='border-left:none'>右</td>
  <td class=xl37>在平</td>
  <td class=xl53>時</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>方式</td>
  <td class=xl53 style='border-left:none'>用停</td>
  <td class=xl53 style='border-left:none'>待售</td>
  <td class=xl37>空</td>
  <td class=xl53>通警</td>
  <td class=xl53 style='border-left:none'>誌標</td>
  <td class=xl53 style='border-left:none'>管制</td>
  <td class=xl53 style='border-left:none'>未依</td>
  <td class=xl53 style='border-left:none'>不將</td>
  <td class=xl53 style='border-left:none'>未依</td>
  <td class=xl53 style='border-left:none'>而逃</td>
  <td class=xl77 style='border-left:none'>規未</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>反</td>
  <td class=xl53 style='border-left:none'>反</td>
  <td class=xl53 style='border-left:none'>放置</td>
  <td class=xl53 style='border-left:none'>路擺</td>
  <td class=xl37>或在車站</td>
  <td class=xl77>第<font class="font10">2</font><font class="font5">款以外</font></td>
  <td class=xl53>或檢</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>廢棄</td>
  <td class=xl120 style='border-left:none'>保管</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>輛</td>
  <td class=xl34>條</td>
  <td class=xl94 style='border-left:none'>優先</td>
  <td class=xl36>行</td>
  <td class=xl53 style='border-left:none'>轉</td>
  <td class=xl37>交道</td>
  <td class=xl53>停</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>車種</td>
  <td class=xl53 style='border-left:none'>車位</td>
  <td class=xl53 style='border-left:none'>或承</td>
  <td class=xl25></td>
  <td class=xl53>察人</td>
  <td class=xl53 style='border-left:none'>線號</td>
  <td class=xl53 style='border-left:none'>規則</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>車輛</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>逸者</td>
  <td class=xl77 style='border-left:none'>列之</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>規</td>
  <td class=xl53 style='border-left:none'>足以</td>
  <td class=xl53 style='border-left:none'>設攤</td>
  <td class=xl37>內休息站</td>
  <td class=xl77>之其他</td>
  <td class=xl53>驗機</td>
  <td class=xl53 style='border-left:none'>拖</td>
  <td class=xl53 style='border-left:none'>車輛</td>
  <td class=xl120 style='border-left:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>與</td>
  <td class=xl34>例</td>
  <td class=xl94 style='border-left:none'>通行</td>
  <td class=xl36>左</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl37>違規</td>
  <td class=xl53>車</td>
  <td class=xl78 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>不依</td>
  <td class=xl53 style='border-left:none'>違規</td>
  <td class=xl53 style='border-left:none'>修之</td>
  <td class=xl25></td>
  <td class=xl53>員取</td>
  <td class=xl53 style='border-left:none'>誌駕</td>
  <td class=xl53 style='border-left:none'>肇事</td>
  <td class=xl53 style='border-left:none'>處理</td>
  <td class=xl53 style='border-left:none'>移置</td>
  <td class=xl53 style='border-left:none'>處置</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl77 style='border-left:none'>條款</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>定</td>
  <td class=xl53 style='border-left:none'>阻礙</td>
  <td class=xl53 style='border-left:none'>位</td>
  <td class=xl37>販賣物品</td>
  <td class=xl77>道路障礙</td>
  <td class=xl53>構採</td>
  <td class=xl53 style='border-left:none'>吊</td>
  <td class=xl53 style='border-left:none'>數量</td>
  <td class=xl120 style='border-left:none'>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>舉</td>
  <td class=xl39>　</td>
  <td class=xl94 style='border-left:none'>　</td>
  <td class=xl36>轉</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>規定</td>
  <td class=xl53 style='border-left:none'>停車</td>
  <td class=xl53 style='border-left:none'>車輛</td>
  <td class=xl25></td>
  <td class=xl53>締</td>
  <td class=xl53 style='border-left:none'>車</td>
  <td class=xl53 style='border-left:none'>致人</td>
  <td class=xl53 style='border-left:none'>而逃</td>
  <td class=xl53 style='border-left:none'>路邊</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl79 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>交通</td>
  <td class=xl78 style='border-left:none'>　</td>
  <td class=xl37>妨礙交通</td>
  <td class=xl77>　</td>
  <td class=xl53>樣測</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl121 style='border-left:none'>35條.12條</td>
  <td class=xl30></td>
  <td class=xl88></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>發</td>
  <td class=xl39>　</td>
  <td class=xl94 style='border-left:none'>　</td>
  <td class=xl36>　</td>
  <td class=xl25></td>
  <td class=xl78>　</td>
  <td class=xl78 style='border-left:none'>　</td>
  <td class=xl80 style='border-left:none'>56條第1項</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl81 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl25></td>
  <td class=xl53>　</td>
  <td class=xl81 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>受傷</td>
  <td class=xl53 style='border-left:none'>逸者</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl79 style='border-left:none'>　</td>
  <td class=xl36>計</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>之物</td>
  <td class=xl78 style='border-left:none'>　</td>
  <td class=xl37></td>
  <td class=xl77>　</td>
  <td class=xl53>定等</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl53 style='border-left:none'>　</td>
  <td class=xl121 style='border-left:none'>3項.57條2</td>
  <td class=xl30></td>
  <td class=xl88></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl129 style='height:15.0pt'>方</td>
  <td class=xl39>　</td>
  <td class=xl94 style='border-left:none'>48條</td>
  <td class=xl91>53<font class="font5">條</font></td>
  <td class=xl69 style='border-left:none'>53<font class="font5">條</font></td>
  <td class=xl53 style='border-left:none'>54條</td>
  <td class=xl53 style='border-left:none'>55條</td>
  <td class=xl53 style='border-left:none'>第1.2.3.4</td>
  <td class=xl53 style='border-left:none'>56條第1項</td>
  <td class=xl53 style='border-left:none'>56條第1項</td>
  <td class=xl53 style='border-left:none'>57條</td>
  <td class=xl53 style='border-left:none'>58條</td>
  <td class=xl53 style='border-left:none'>60條</td>
  <td class=xl53 style='border-left:none'>60條第2項</td>
  <td class=xl53 style='border-left:none'>61條</td>
  <td class=xl53 style='border-left:none'>62條</td>
  <td class=xl53 style='border-left:none'>62條</td>
  <td class=xl53 style='border-left:none'>62條</td>
  <td class=xl37>62條</td>
  <td class=xl77>12條至</td>
  <td class=xl36>　</td>
  <td class=xl53 style='border-left:none'>69條</td>
  <td class=xl53 style='border-left:none'>78條至</td>
  <td class=xl53 style='border-left:none'>82條第1項</td>
  <td class=xl53 style='border-left:none'>82條第1項</td>
  <td class=xl37>83條</td>
  <td class=xl77>82條</td>
  <td class=xl36 x:str="'35條">35條</td>
  <td class=xl53 style='border-left:none'>56條</td>
  <td class=xl53 style='border-left:none'>82條之1</td>
  <td class=xl121 style='border-left:none'>項.62條6項</td>
  <td class=xl30></td>
  <td class=xl88></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl130 style='height:15.0pt'>式</td>
  <td class=xl40>　</td>
  <td class=xl95 style='border-left:none'>第<font class="font10">2</font><font
  class="font5">項</font></td>
  <td class=xl42>第1項</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">2</font><font
  class="font5">項</font></td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl56>　</td>
  <td class=xl54>5.<font class="font10">6.</font><font class="font5">7.8款</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">9</font><font
  class="font5">款</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">10</font><font
  class="font5">款</font></td>
  <td class=xl55 style='border-left:none'>　</td>
  <td class=xl57>第<font class="font10">3</font><font class="font5">款</font></td>
  <td class=xl54>第<font class="font10">1</font><font class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font5">3款</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl57>第1項</td>
  <td class=xl54>第<font class="font10">2</font><font class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">4</font><font
  class="font5">項</font></td>
  <td class=xl82 style='border-left:none'>62條</td>
  <td class=xl58>　</td>
  <td class=xl54 style='border-left:none'>至76條</td>
  <td class=xl54 style='border-left:none'>81條之1</td>
  <td class=xl54 style='border-left:none'>第1款</td>
  <td class=xl54 style='border-left:none'>第<font class="font10">10</font><font
  class="font5">款</font></td>
  <td class=xl57>　</td>
  <td class=xl82>至84條</td>
  <td class=xl42>第<font class="font10">5</font><font class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第<font class="font10">3</font><font
  class="font5">項</font></td>
  <td class=xl54 style='border-left:none'>第1項</td>
  <td class=xl122 style='border-left:none'>85條之2</td>
  <td class=xl30></td>
  <td class=xl88></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl144 style='border-right:1.0pt solid black;
  height:15.0pt'>合計</td>
  <td class=xl45 align=right x:num x:fmla="=SUM(C61,C64,C67,C70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(D61,D64,D67,D70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(E61,E64,E67,E70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(F61,F64,F67,F70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(G61,G64,G67,G70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(H61,H64,H67,H70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(I61,I64,I67,I70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(J61,J64,J67,J70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(K61,K64,K67,K70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(L61,L64,L67,L70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(M61,M64,M67,M70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(N61,N64,N67,N70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(O61,O64,O67,O70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(P61,P64,P67,P70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(Q61,Q64,Q67,Q70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(R61,R64,R67,R70)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num x:fmla="=SUM(S61,S64,S67,S70)">0</td>
  <td class=xl109 align=right style='border-top:none;border-left:none' x:num x:fmla="=SUM(T61,T64,T67,T70)">0</td>
  <td class=xl44 align=right x:num x:fmla="=SUM(V60:AA60)">0</td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("69,70,71,72,73,74,75,76","1")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("78,79,80,81","1")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("8210101","5")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("8211001","5")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("8310101","5")%></td>
  <td class=xl109 align=right style='border-top:none;border-left:none' x:num   x:fmla="=<%=GetValue2("82,83,84","1")%>-X60-Y60-Z60"></td>
  <td class=xl44 align=right x:num><%=GetValue2("355","2")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("563","2")%></td>
  <td class=xl45 align=right style='border-left:none' x:num><%=GetValue2("821-1","7")%></td>
  <td class=xl83 align=right style='border-left:none' x:num x:fmla="=SUM(AE62:AE63,AE65:AE66,AE68:AE70)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl45 align=right style='border-top:none' x:num  x:fmla="=SUM(C62:C63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(D62:D63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(E62:E63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(F62:F63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(G62:G63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(H62:H63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(I62:I63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(J62:J63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(K62:K63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(L62:L63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(M62:M63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(N62:N63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(O62:O63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(P62:P63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Q62:Q63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(R62:R63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(S62:S63)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(T62:T63)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(U62:U63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(V62:V63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(W62:W63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(X62:X63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Y62:Y63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Z62:Z63)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AA62:AA63)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(AB62:AB63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AC62:AC63)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AD62:AD63)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AE62:AE63)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>汽車</td>
  <td class=xl43 style='border-top:none;border-left:none'>逕舉</td>
  <td class=xl45 style='border-top:none' x:num><%=C62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S62%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=GetTotalValue("1","2")%>-SUM(E18:AE18,C40:AE40,C62:S62)<1,0,<%=GetTotalValue("1","2")%>-SUM(E18:AE18,C40:AE40,C62:S62))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V62:AA62)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("1","2","35","1"))+cdbl(GetValue3("1","2","123","2"))+cdbl(GetValue3("1","2","572","2"))+cdbl(GetValue3("1","2","626","2"))+cdbl(GetValue3("1","2","85-2","6"))%>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>　</td>
  <td class=xl43 style='border-top:none;border-left:none'>攔停</td>
  <td class=xl45 style='border-top:none' x:num><%=C63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R63%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S63%></td>
  <td class=xl85 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=CDbl(GetTotalValue("1","1"))+CDbl(GetTotalValue2)%>-SUM(E19:AE19,C41:AE41,C63:S63)<1,0,<%=CDbl(GetTotalValue("1","1"))+CDbl(GetTotalValue2)%>-SUM(E19:AE19,C41:AE41,C63:S63))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V63:AA63)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("1","1","35","1"))+cdbl(GetValue3("1","1","123","2"))+cdbl(GetValue3("1","1","572","2"))+cdbl(GetValue3("1","1","626","2"))+cdbl(GetValue3("1","1","85-2","6"))%>　</td>  
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl93 style='height:15.0pt;border-top:none'>550cc以上</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl45 align=right style='border-top:none' x:num  x:fmla="=SUM(C65:C66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(D65:D66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(E65:E66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(F65:F66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(G65:G66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(H65:H66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(I65:I66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(J65:J66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(K65:K66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(L65:L66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(M65:M66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(N65:N66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(O65:O66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(P65:P66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Q65:Q66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(R65:R66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(S65:S66)">0</td>
  <td class=xl83 align=right style='border-left:none' x:num  x:fmla="=SUM(T65:T66)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(U65:U66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(V65:V66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(W65:W66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(X65:X66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Y65:Y66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Z65:Z66)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AA65:AA66)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(AB65:AB66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AC65:AC66)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AD65:AD66)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AE65:AE66)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl94 style='height:15.0pt'>大型重</td>
  <td class=xl43 style='border-top:none;border-left:none'>逕舉</td>
<td class=xl45 style='border-top:none' x:num><%=C65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S65%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=GetTotalValue("4","2")%>-SUM(E21:AE21,C43:AE43,C65:S65)<1,0,<%=GetTotalValue("4","2")%>-SUM(E21:AE21,C43:AE43,C65:S65))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V65:AA65)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("4","2","35","1"))+cdbl(GetValue3("4","2","123","2"))+cdbl(GetValue3("4","2","572","2"))+cdbl(GetValue3("4","2","626","2"))+cdbl(GetValue3("4","2","85-2","6"))%>　</td>  
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl131 style='height:15.0pt'>型機車</td>
  <td class=xl43 style='border-top:none;border-left:none'>攔停</td>
<td class=xl45 style='border-top:none' x:num><%=C66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S66%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=GetTotalValue("4","1")%>-SUM(E22:AE22,C44:AE44,C66:S66)<1,0,<%=GetTotalValue("4","1")%>-SUM(E22:AE22,C44:AE44,C66:S66))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V66:AA66)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("4","1","35","1"))+cdbl(GetValue3("4","1","123","2"))+cdbl(GetValue3("4","1","572","2"))+cdbl(GetValue3("4","1","626","2"))+cdbl(GetValue3("4","1","85-2","6"))%>　</td>  
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td rowspan=3 height=60 class=xl138 width=85 style='border-bottom:.5pt solid black;
  height:45.0pt;border-top:none;width:64pt'>未滿<br>
    550cc<br>
    機車</td>
  <td class=xl43 style='border-top:none;border-left:none'>小計</td>
  <td class=xl45 align=right style='border-top:none' x:num  x:fmla="=SUM(C68:C69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(D68:D69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(E68:E69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(F68:F69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(G68:G69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(H68:H69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(I68:I69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(J68:J69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(K68:K69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(L68:L69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(M68:M69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(N68:N69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(O68:O69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(P68:P69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Q68:Q69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(R68:R69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(S68:S69)">0</td>
  <td class=xl83 align=right style='border-left:none' x:num  x:fmla="=SUM(T68:T69)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(U68:U69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(V68:V69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(W68:W69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(X68:X69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Y68:Y69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(Z68:Z69)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AA68:AA69)">0</td>
  <td class=xl44 align=right style='border-top:none' x:num  x:fmla="=SUM(AB68:AB69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AC68:AC69)">0</td>
  <td class=xl45 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AD68:AD69)">0</td>
  <td class=xl83 align=right style='border-top:none;border-left:none' x:num  x:fmla="=SUM(AE68:AE69)">0</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>逕舉</td>
  <td class=xl45 style='border-top:none' x:num><%=C68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S68%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=GetTotalValue("2","2")%>-SUM(E24:AE24,C46:AE46,C68:S68)<1,0,<%=GetTotalValue("2","2")%>-SUM(E24:AE24,C46:AE46,C68:S68))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V68:AA68)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("2","2","35","1"))+cdbl(GetValue3("2","2","123","2"))+cdbl(GetValue3("2","2","572","2"))+cdbl(GetValue3("2","2","626","2"))+cdbl(GetValue3("2","2","85-2","6"))%>　</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl68 style='height:15.0pt;border-top:none'>攔停</td>
  <td class=xl45 style='border-top:none' x:num><%=C69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=D69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=E69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=F69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=G69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=H69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=I69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=J69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=K69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=L69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=M69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=N69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=O69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=P69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=Q69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=R69%></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num><%=S69%></td>
  <td class=xl85 style='border-top:none;border-left:none' x:num x:fmla="=if(<%=GetTotalValue("2","1")%>-SUM(E25:AE25,C47:AE47,C69:S69)<1,0,<%=GetTotalValue("2","1")%>-SUM(E25:AE25,C47:AE47,C69:S69))"></td>
  <td class=xl44 style='border-top:none' x:numx  x:fmla="=Sum(V69:AA69)"></td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl125 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl44 style='border-top:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl45 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl123 style='border-top:none' x:num><%=cdbl(GetValue3("2","1","35","1"))+cdbl(GetValue3("2","1","123","2"))+cdbl(GetValue3("2","1","572","2"))+cdbl(GetValue3("2","1","626","2"))+cdbl(GetValue3("2","1","85-2","6"))%>　</td>  
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl132 style='height:15.0pt;border-top:none'>動力機械</td>
  <td class=xl46 style='border-top:none;border-left:none'>小計</td>
  <td class=xl59 style='border-top:none' x:num><%=C70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=D70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=E70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=F70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=G70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=H70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=I70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=J70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=K70%></td>
  <td class=xl59 style='border-left:none' x:num><%=L70%></td>
  <td class=xl59 style='border-left:none' x:num><%=M70%></td>
  <td class=xl59 style='border-left:none' x:num><%=N70%></td>
  <td class=xl59 style='border-left:none' x:num><%=O70%></td>
  <td class=xl59 style='border-top:none;border-left:none' x:num><%=P70%></td>
  <td class=xl59 style='border-left:none' x:num><%=Q70%></td>
  <td class=xl59 style='border-left:none' x:num><%=R70%></td>
  <td class=xl59 style='border-left:none' x:num><%=S70%></td>
  <td class=xl86 style='border-left:none' x:numx: x:fmla="=if(<%=GetTotalValue("5","1")%>-SUM(E26:AE26,C48:AE48,C70:S70)<1,0,<%=GetTotalValue("5","1")%>-SUM(E26:AE26,C48:AE48,C70:S70))"></td>
  <td class=xl92 style='border-top:none' x:num>0</td>
  <td class=xl59 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl59 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl59 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl59 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl59 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl86 style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl47 align=right style='border-top:none' x:num>0</td>
  <td class=xl48 align=right style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl48 align=right style='border-top:none;border-left:none' x:num>0</td>
  <td class=xl124 align=right style='border-top:none' x:num><%=cdbl(GetValue3("5","1","35","1"))+cdbl(GetValue3("5","1","123","2"))+cdbl(GetValue3("5","1","572","2"))+cdbl(GetValue3("5","1","626","2"))+cdbl(GetValue3("5","1","85-2","6"))%></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl25 style='height:15.0pt'>填<span
  style='mso-spacerun:yes'>&nbsp; </span>表</td>
  <td class=xl37></td>
  <td class=xl60></td>
  <td class=xl30></td>
  <td class=xl60></td>
  <td class=xl25></td>
  <td class=xl25>審<span style='mso-spacerun:yes'>&nbsp; </span>核</td>
  <td colspan=3 class=xl60 style='mso-ignore:colspan'></td>
  <td class=xl30></td>
  <td class=xl25></td>
  <td class=xl25 colspan=2 style='mso-ignore:colspan'>主辦業務人員</td>
  <td colspan=2 class=xl60 style='mso-ignore:colspan'></td>
  <td class=xl25></td>
  <td colspan=2 class=xl60 style='mso-ignore:colspan'></td>
  <td class=xl25>機關長官</td>
  <td colspan=6 class=xl60 style='mso-ignore:colspan'></td>
    <%
    DBDate = ""
    sql = "select sysdate from Dual"
    Set RSSystem = Conn.Execute(sql)
    DBDate = RSSystem("sysdate")
	Set RSSystem=nothing
  %>
  <td colspan=5 class=xl75>中華民國<%=CInt(Mid(DBDate,1,4))-1911%>年<%=datepart("m",DBDate)%>月<%=datepart("d",DBDate)%>日編製</td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr class=xl30 height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 colspan=2 class=xl37 style='height:15.0pt;mso-ignore:colspan'></td>
  <td colspan=4 class=xl60 style='mso-ignore:colspan'></td>
  <td class=xl30></td>
  <td colspan=4 class=xl60 style='mso-ignore:colspan'></td>
  <td class=xl25></td>
  <td class=xl25 colspan=2 style='mso-ignore:colspan'>主辦統計人員</td>
  <td colspan=17 class=xl60 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl30 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=21 style='mso-height-source:userset;height:15.95pt'>
  <td height=21 class=xl25 colspan=17 style='height:15.95pt;mso-ignore:colspan'>資料來源：各分局(金門縣、連江縣為警察所)、專業警察機關(保安警察第二總隊、航空警察局、國家公園警察大隊、國道公路警察局、基隆、臺中、高雄、花蓮港務警察局、鐵路警察局)。</td>
  <td colspan=13 class=xl25 style='mso-ignore:colspan'></td>
  <td class=xl60></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl25 colspan=18 style='height:16.5pt;mso-ignore:colspan'>填表說明：(一)本表編製1式2份，先送會計室(統計室)會核，並經機關長官核章後，1份送會計室(統計室)，1份自存外，本表於規定期限內由網際網路線上傳送至內政部警政署警政統計資料庫。</td>
  <td colspan=13 class=xl25 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl25 colspan=27 style='height:16.5pt;mso-ignore:colspan'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>(二)「慢車違反規定」、「行人違反規定」、「在道路上堆積放置足以妨礙交通之物」、「未經許可在在道路道路擺設攤位」、「在車道交通島散發廣告物等或在車站內休息站販賣物品妨礙交通」、「82條第1項第1.10款、83條第1項第2款以外之其他道路障礙」等6欄僅填舉發件數。</td>
  <td colspan=4 class=xl25 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl25 colspan=16 style='height:16.5pt;mso-ignore:colspan'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>(三)「強制移由醫療或檢驗機構採樣測定」、「違規停車拖吊」、「查報佔用道路廢棄車輛數量」、「違規車輛移置保管」等4欄可填汽車、機車及動力機械舉發件數。</td>
  <td colspan=8 class=xl25 style='mso-ignore:colspan'></td>
  <td class=xl61></td>
  <td class=xl37></td>
  <td colspan=4 class=xl25 style='mso-ignore:colspan'></td>
  <td class=xl37></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 class=xl25 colspan=11 style='height:16.5pt;mso-ignore:colspan'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>(四)舉發總件數：(移公路監理機關舉發件數)+(警察機關舉發件數)。<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>(五)本表自97年1月1日起實施。</td>
  <td colspan=13 class=xl25 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl37 style='mso-ignore:colspan'></td>
  <td colspan=5 class=xl25 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 colspan=31 class=xl25 style='height:16.5pt;mso-ignore:colspan'></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=22 style='height:16.5pt'>
  <td height=22 colspan=9 class=xl25 style='height:16.5pt;mso-ignore:colspan'></td>
  <td class=xl61></td>
  <td colspan=21 class=xl25 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl61 style='mso-ignore:colspan'></td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=85 style='width:64pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=68 style='width:51pt'></td>
 </tr>
 <![endif]>
</table>

</body>

</html>
<%
fYear=year(now)-1911
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay


	If Sys_City="台中市" Then
		If InStr(Sys_UnitID,"0460")>0 And InStr(Sys_UnitID,"0406")>0 Then
			fname=Sys_City&"_"&Mid(Trim(Request("startDate_q")),1,Len(Trim(Request("startDate_q")))-2)&"_舉發違反道路交通管理事件成果表.xls"
		Else
			fname=fYear&fMnoth & fDay & filename & "_舉發違反道路交通管理事件成果表.xls"
		End if
	else
		fname=Sys_City&"_"&Mid(Trim(Request("startDate_q")),1,Len(Trim(Request("startDate_q")))-2)&"_舉發違反道路交通管理事件成果表.xls"
	End If 


Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>