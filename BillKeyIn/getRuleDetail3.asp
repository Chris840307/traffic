<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getRuleDetail.asp
	'違規事實
RuleOrder=trim(request("RuleOrder"))
' Or trim(request("RuleID"))="2710001" Or trim(request("RuleID"))="2710002" Or trim(request("RuleID"))="2710003" Or trim(request("RuleID"))="2710004" Or trim(request("RuleID"))="2710005" Or trim(request("RuleID"))="2710006" Or trim(request("RuleID"))="2710007" Or trim(request("RuleID"))="2710008" Or trim(request("RuleID"))="2720001" Or trim(request("RuleID"))="2720002" Or trim(request("RuleID"))="2720003" Or trim(request("RuleID"))="2720004" Or trim(request("RuleID"))="4000001" Or trim(request("RuleID"))="4000002" Or trim(request("RuleID"))="4000003" Or trim(request("RuleID"))="4310201" Or trim(request("RuleID"))="4310202" Or trim(request("RuleID"))="4310203" Or trim(request("RuleID"))="4310204" Or trim(request("RuleID"))="4310205" Or trim(request("RuleID"))="4310206" Or trim(request("RuleID"))="4310207"  Or trim(request("RuleID"))="4310208" Or trim(request("RuleID"))="4310209" Or trim(request("RuleID"))="5400101" Or trim(request("RuleID"))="5400102" Or trim(request("RuleID"))="5400103" Or trim(request("RuleID"))="5400104" Or trim(request("RuleID"))="5400105" Or trim(request("RuleID"))="5400106" Or trim(request("RuleID"))="5400107"  Or trim(request("RuleID"))="5400108" Or trim(request("RuleID"))="5400109" Or trim(request("RuleID"))="5400201" Or trim(request("RuleID"))="5400202" Or trim(request("RuleID"))="5400203" Or trim(request("RuleID"))="5400204" Or trim(request("RuleID"))="5400205" Or trim(request("RuleID"))="5400206" Or trim(request("RuleID"))="5400301" Or trim(request("RuleID"))="5400302" Or trim(request("RuleID"))="5400303" Or trim(request("RuleID"))="5400304" Or trim(request("RuleID"))="5400305" Or trim(request("RuleID"))="5400306" Or trim(request("RuleID"))="5400307" Or trim(request("RuleID"))="5400308" Or trim(request("RuleID"))="5400309" Or trim(request("RuleID"))="5400310" Or trim(request("RuleID"))="5400311" Or trim(request("RuleID"))="5400312" Or trim(request("RuleID"))="5400313" Or trim(request("RuleID"))="5400314" Or trim(request("RuleID"))="5400315" Or trim(request("RuleID"))="4511401" Or trim(request("RuleID"))="4511402" Or trim(request("RuleID"))="4511403" Or trim(request("RuleID"))="4511404"
if trim(request("RuleID"))="1210403" or trim(request("RuleID"))="1210404" or trim(request("RuleID"))="1210503" or trim(request("RuleID"))="1210504" or trim(request("RuleID"))="1210602" or trim(request("RuleID"))="1210806" or trim(request("RuleID"))="1210807" or trim(request("RuleID"))="1210808" or trim(request("RuleID"))="1210809" or trim(request("RuleID"))="1210800" or trim(request("RuleID"))="3140011" or trim(request("RuleID"))="3140012" or trim(request("RuleID"))="4200002" or trim(request("RuleID"))="1210902" or trim(request("RuleID"))="3311105" or trim(request("RuleID"))="3311106" or trim(request("RuleID"))="3010201" or trim(request("RuleID"))="3010202" or trim(request("RuleID"))="3010203" or trim(request("RuleID"))="3010204" or trim(request("RuleID"))="3010205" or trim(request("RuleID"))="3010206" or trim(request("RuleID"))="4340011" or trim(request("RuleID"))="4500101" or trim(request("RuleID"))="31100011" or trim(request("RuleID"))="31100021" or trim(request("RuleID"))="31200011" or trim(request("RuleID"))="31200021" or trim(request("RuleID"))="3510116" or trim(request("RuleID"))="3510117" or trim(request("RuleID"))="3510118" or trim(request("RuleID"))="3510119" Or Left(trim(request("RuleID")),3)="450" Or Left(trim(request("RuleID")),3)="571" then
	if RuleOrder="1" then 
%>
		TDLawErrorLog1=1;
<%
	elseif RuleOrder="2" then
%>
		TDLawErrorLog2=1;
<%	
	elseif RuleOrder="3" then	
%>
		TDLawErrorLog3=1;
<%		
	elseif RuleOrder="4" then
%>
		TDLawErrorLog4=1;
<%		
	end if
end if

if trim(request("RuleID"))="1210403" then
%>
	alert("對應新條款1210401，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210404" then
%>
	alert("對應新條款1210402，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210503" then
%>
	alert("對應新條款1210501，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210504" then
%>
	alert("對應新條款1210502，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210602" then
%>
	alert("對應新條款1210601，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210806" then
%>
	alert("對應新條款1210801，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210807" then
%>
	alert("對應新條款1210802，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210808" then
%>
	alert("對應新條款1210803，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210809" then
%>
	alert("對應新條款1210804，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210800" then
%>
	alert("對應新條款1210805，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="1210902" then
%>
	alert("對應新條款1210901，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3140011" then
%>
	alert("對應新條款3140001，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3140012" then
%>
	alert("對應新條款3140002，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="4200002" then
%>
	alert("對應新條款4200001，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3311105" Or trim(request("RuleID"))="3311106" then
%>
	alert("對應新條款 3010207 ~ 3010212，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010201" then
%>
	alert("對應新條款3010207，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010202" then
%>
	alert("對應新條款3010208，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010203" then
%>
	alert("對應新條款3010209，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010204" then
%>
	alert("對應新條款3010210，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010205" then
%>
	alert("對應新條款3010211，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3010206" then
%>
	alert("對應新條款3010212，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="31100011" Or trim(request("RuleID"))="31100021" then
%>
	alert("對應新條款 31100031 ~ 31100141，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="31200011" Or trim(request("RuleID"))="31200021" then
%>
	alert("對應新條款 31200031 ~ 31200141，入案請鍵入新條款。");
<%
elseif trim(request("RuleID"))="3510116" Or trim(request("RuleID"))="3510117" Or trim(request("RuleID"))="3510118" Or trim(request("RuleID"))="3510119" then
%>
	alert("對應新條款 3510134、3510137 ~ 3510139，入案請鍵入新條款。");
<%
ElseIf Left(trim(request("RuleID")),3)="450" Then
%>
	alert("法條4500101 ~ 4501506已停用,對應新條款4510101 ~ 4511506，入案請鍵入新條款。");
<%	
ElseIf Left(trim(request("RuleID")),3)="571" Then
%>
	alert("法條5710101 ~ 5710106已停用,對應新條款5700106 ~ 5700106，入案請鍵入新條款。");
<%
ElseIf trim(request("RuleID"))="4340011" Then
%>
	alert("法條4340011已停用。");
<%
ElseIf rulechange1015="Yes" Then
'trim(request("RuleID"))="2710001" Or trim(request("RuleID"))="2710002" Or trim(request("RuleID"))="2710003" Or trim(request("RuleID"))="2710004" Or trim(request("RuleID"))="2710005" Or trim(request("RuleID"))="2710006" Or trim(request("RuleID"))="2710007" Or trim(request("RuleID"))="2710008" Or trim(request("RuleID"))="2720001" Or trim(request("RuleID"))="2720002" Or trim(request("RuleID"))="2720003" Or trim(request("RuleID"))="2720004"
%>
	alert("2012年10月15日開始，此代碼不適用。");
<%
ElseIf rulechange1015="Yes" Then
'trim(request("RuleID"))="4000001" Or trim(request("RuleID"))="4000002" Or trim(request("RuleID"))="4000003"
%>
	alert("4000001 ~ 4000003已停用,對應新條款4000005 ~ 4000007，入案請鍵入新條款。");
<%
ElseIf rulechange1015="Yes" Then
'trim(request("RuleID"))="4310201" Or trim(request("RuleID"))="4310202" Or trim(request("RuleID"))="4310203" Or trim(request("RuleID"))="4310204" Or trim(request("RuleID"))="4310205" Or trim(request("RuleID"))="4310206" Or trim(request("RuleID"))="4310207"  Or trim(request("RuleID"))="4310208" Or trim(request("RuleID"))="4310209"
%>
	alert("4310201 ~ 4310209已停用,對應新條款4310210 ~ 4310218，入案請鍵入新條款。");
<%
ElseIf rulechange1015="Yes" Then
'trim(request("RuleID"))="5400101" Or trim(request("RuleID"))="5400102" Or trim(request("RuleID"))="5400103" Or trim(request("RuleID"))="5400104" Or trim(request("RuleID"))="5400105" Or trim(request("RuleID"))="5400106" Or trim(request("RuleID"))="5400107"  Or trim(request("RuleID"))="5400108" Or trim(request("RuleID"))="5400109"
%>
	alert("5400101 ~ 5400109已停用,對應新條款5400110 ~ 5400118，入案請鍵入新條款。");
<%
ElseIf  rulechange1015="Yes" Then
'trim(request("RuleID"))="4511401" Or trim(request("RuleID"))="4511402" Or trim(request("RuleID"))="4511403" Or trim(request("RuleID"))="4511404" 
%>
	alert("4511401 ~ 4511404已停用,對應新條款4511405 ~ 4511408，入案請鍵入新條款。");
<%
ElseIf rulechange1015="Yes" Then
' trim(request("RuleID"))="5400201" Or trim(request("RuleID"))="5400202" Or trim(request("RuleID"))="5400203" Or trim(request("RuleID"))="5400204" Or trim(request("RuleID"))="5400205" Or trim(request("RuleID"))="5400206" 
%>
	alert("5400201 ~ 5400206已停用,對應新條款5400207 ~ 5400212，入案請鍵入新條款。");
<%
ElseIf rulechange1015="Yes" Then
' trim(request("RuleID"))="5400301" Or trim(request("RuleID"))="5400302" Or trim(request("RuleID"))="5400303" Or trim(request("RuleID"))="5400304" Or trim(request("RuleID"))="5400305" Or trim(request("RuleID"))="5400306" Or trim(request("RuleID"))="5400307" Or trim(request("RuleID"))="5400308" Or trim(request("RuleID"))="5400309" Or trim(request("RuleID"))="5400310" Or trim(request("RuleID"))="5400311" Or trim(request("RuleID"))="5400312" Or trim(request("RuleID"))="5400313" Or trim(request("RuleID"))="5400314" Or trim(request("RuleID"))="5400315"
%>
	alert("5400301 ~ 5400315已停用,對應新條款5400316 ~ 5400330，入案請鍵入新條款。");
<%
else
	
	if trim(request("CarSimpleID"))="" then
		CarSimple=0
	else
		CarSimple=trim(request("CarSimpleID"))
	end If
	response.write left(trim(request("RuleID")),4)
	if left(trim(request("RuleID")),4)="2110" or trim(request("RuleID"))="4310102" or trim(request("RuleID"))="4310103" Or left(trim(request("RuleID")),4)="2210" then
		if CarSimple=1 or CarSimple=2 then
			strCarImple=" and CarSimpleID in ('5','0')"
		elseif CarSimple="3" or CarSimple="4" then
			strCarImple=" and CarSimpleID in ('3','0')"
		end if
	end If
	
	theRuleVer=trim(request("RuleVer"))
	IllArrayStr=""
	level1ArrayStr=""
	SimpleIDArrayStr=""
	strLaw="select * from Law where ItemID='"&trim(request("RuleID"))&"' and Version='"&theRuleVer&"'"&strCarImple&" order by CarSimpleID Desc"
	response.write "<br>"&strlaw
	set rsLaw=conn.execute(strLaw)
	if not rsLaw.Bof then 
		if SimpleIDArrayStr="" then
			SimpleIDArrayStr=trim(rsLaw("CarSimpleID"))
		else
			SimpleIDArrayStr=SimpleIDArrayStr&","&trim(rsLaw("CarSimpleID"))
		end if
		if IllArrayStr="" then
			IllArrayStr=trim(rsLaw("IllegalRule"))
		else
			IllArrayStr=IllArrayStr&","&trim(rsLaw("IllegalRule"))
		end if
		if level1ArrayStr="" then
			level1ArrayStr=trim(rsLaw("Level1"))
		else
			level1ArrayStr=level1ArrayStr&","&trim(rsLaw("Level1"))
		end if
	end if
	rsLaw.close
	set rsLaw=nothing
%>setLawDetail("<%=RuleOrder%>","<%=IllArrayStr%>","<%=level1ArrayStr%>");<%
end if

conn.close
set conn=nothing
%>
