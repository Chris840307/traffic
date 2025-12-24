<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
' 檔案名稱： getChkCarSimpleIDandRule.asp
	'檢查法條跟車種相不相符

		'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	theCarSimpleID=trim(request("CarSimpleID"))
	'法條一
	Rule1Flag=0
	FlagRuleDetail=0
	strRule1=""
	if trim(request("IllRule1"))<>"" then
		strRule1=" and Rule1='"&trim(request("IllRule1"))&"'"
		Rule1Flag=1
		'檢查車種跟法條內容相不相符
		if left(request("IllRule1"),2)="31" then
			strRuleDetail1="select IllegalRule from Law where ItemID='"&trim(request("IllRule1"))&"'"
			set rsRuleDetail1=conn.execute(strRuleDetail1)
			If Not rsRuleDetail1.eof Then
				if InStr(rsRuleDetail1("IllegalRule"),"機器腳踏車")>0 and (theCarSimpleID="1" or theCarSimpleID="2") then
					FlagRuleDetail=1
				elseif (InStr(rsRuleDetail1("IllegalRule"),"小客車")>0 or InStr(rsRuleDetail1("IllegalRule"),"汽車")>0) and (theCarSimpleID="3" or theCarSimpleID="4") then
					FlagRuleDetail=1
				end if
			end if
			rsRuleDetail1.close
			set rsRuleDetail1=nothing
		end if
	end if

	'法條二
	Rule2Flag=0
	strRule2=""
	if trim(request("IllRule2"))<>"" then
		strRule2=" and Rule2='"&trim(request("IllRule2"))&"'"
		Rule2Flag=1
		'檢查車種跟法條內容相不相符
		if left(request("IllRule2"),2)="31" then
			strRuleDetail2="select IllegalRule from Law where ItemID='"&trim(request("IllRule2"))&"'"
			set rsRuleDetail2=conn.execute(strRuleDetail2)
			If Not rsRuleDetail2.eof Then
				if InStr(rsRuleDetail2("IllegalRule"),"機器腳踏車")>0 and (theCarSimpleID="1" or theCarSimpleID="2") then
					FlagRuleDetail=1
				elseif (InStr(rsRuleDetail2("IllegalRule"),"小客車")>0 or InStr(rsRuleDetail2("IllegalRule"),"汽車")>0) and (theCarSimpleID="3" or theCarSimpleID="4") then
					FlagRuleDetail=1
				end if
			end if
			rsRuleDetail2.close
			set rsRuleDetail2=nothing
		end if
	end if

%>setChkCarSimpleIDandRule("<%=FlagRuleDetail%>");
<%
conn.close
set conn=nothing
%>
function setChkCarSimpleIDandRule(RuleDetail){
	if (RuleDetail==1){
		if(confirm('違規事實與簡式車種不符，請確認是否正確。\n是否確定要存檔？')){
			document.myForm.kinds.value="DB_insert";
			document.myForm.submit();
		}
	}else{
		document.myForm.kinds.value="DB_insert";
		document.myForm.submit();
	}
}
