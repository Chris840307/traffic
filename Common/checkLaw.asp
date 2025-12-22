<%
showBarCode=true
if left(trim(Sys_Rule1),2)="40" or left(trim(Sys_Rule1),5)="33101" then
	if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
		if Sys_IllegalSpeed-Sys_RuleSpeed>60 then showBarCode=false
	end If 
elseif left(trim(Sys_Rule1),2)="12" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="13" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="15102" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="15105" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="16105" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="17" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="18" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="20" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="21" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="23" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="24" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="25103" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="26" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="272" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="293" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="29" and (len(Sys_Rule1)=8 and right(Sys_Rule1,1)="2") then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="294" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="303" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="314" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="34" and cdbl(right(Sys_Rule1,2))>2 then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="35" and left(trim(Sys_Rule1),3)<>"358" and not (len(Sys_Rule1)=8 and right(Sys_Rule1,1)="1") then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="362" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="363" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="37" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="43" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="452" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="453" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="45011" then
	showBarCode=false
elseif left(trim(Sys_Rule1),5)="45111" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="54" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="601" then
	showBarCode=false
elseif left(trim(Sys_Rule1),2)="61" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="621" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="624" then
	showBarCode=false
elseif left(trim(Sys_Rule1),3)="625" then
	showBarCode=false
end if
If showBarCode=true and Sys_Rule2<>"" Then
	if left(trim(Sys_Rule2),2)="40" or left(trim(Sys_Rule2),5)="33101" then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			if Sys_IllegalSpeed-Sys_RuleSpeed>60 then showBarCode=false
		end If 
	elseif left(trim(Sys_Rule2),2)="12" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="13" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="15102" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="15105" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="16105" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="17" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="18" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="20" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="21" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="23" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="24" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="25103" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="26" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="272" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="293" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="29" and (len(Sys_Rule2)=8 and right(Sys_Rule2,1)="2") then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="294" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="303" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="314" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="34" and cdbl(right(Sys_Rule2,2))>2 then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="35" and left(trim(Sys_Rule2),3)<>"358" and not (len(Sys_Rule2)=8 and right(Sys_Rule2,1)="1") then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="362" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="363" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="37" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="43" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="452" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="453" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="45011" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),5)="45111" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="54" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="601" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),2)="61" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="621" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="624" then
		showBarCode=false
	elseif left(trim(Sys_Rule2),3)="625" then
		showBarCode=false
	end if
End if
%>