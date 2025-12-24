<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getRuleDetail.asp
	'違規事實
	RuleOrder=trim(request("RuleOrder"))
	if trim(request("CarSimpleID"))="" then
		CarSimple=0
	else
		CarSimple=trim(request("CarSimpleID"))
	end If
	strCarImple=""
				if trim(request("CarSimpleID"))="1" then
					strCarImple=" and CarSimpleID in ('5','4')"
				elseif trim(request("CarSimpleID"))="2" then
					strCarImple=" and CarSimpleID in ('5','4')"
				elseif trim(request("CarSimpleID"))="3" then
					strCarImple=" and CarSimpleID in ('3')"
				elseif trim(request("CarSimpleID"))="4" then
					strCarImple=" and CarSimpleID in ('3')"
				end If

				
	if left(trim(request("RuleID")),4)="2110" or trim(request("RuleID"))="4310102" or trim(request("RuleID"))="4310103" then
		if CarSimple=1 or CarSimple=2 then
			strCarImple=" and CarSimpleID in ('5','0')"
		elseif CarSimple=3 or CarSimple=4 then
			strCarImple=" and CarSimpleID in ('3','0')"
		end if
	end if
	theRuleVer=trim(request("RuleVer"))
	IllArrayStr=""
	level1ArrayStr=""
	SimpleIDArrayStr=""
	strLaw="select * from Law where ItemID='"&trim(request("RuleID"))&"' and Version='"&theRuleVer&"'"&strCarImple&" order by CarSimpleID Desc"
	set rsLaw=conn.execute(strLaw)
	If rsLaw.Bof Then
		strLaw="select * from Law where ItemID='"&trim(request("RuleID"))&"' and Version='"&theRuleVer&"' and CarSimpleID in ('0') order by CarSimpleID Desc"
		set rsLaw=conn.execute(strLaw)
	End If
	
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

%>
	

		if ("<%=IllArrayStr%>" != ""){

			Layer1.innerHTML="<%=IllArrayStr%>"
			myForm.ForFeit1.value=<%=level1ArrayStr%>;
			TDLawErrorLog1=0;
			myForm.ForFeit1.select();

		}else{
			Layer1.innerHTML=" ";
			myForm.ForFeit1.value="";
			TDLawErrorLog1=1;
		}
	



<%
conn.close
set conn=nothing
%>
