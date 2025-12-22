<%
if trim(Session("FuncID"))="" or trim(Session("FuncID"))="" then
	Response.Redirect "/traffic/Traffic_Login.asp?Error=叫祅セ╰参"

end if
FuncIDtemp=trim(Session("FuncID"))
'浪琩琌Τㄏノセ╰参ぇ舦
public function AuthorityCheck(FID)
	FunctionTemp=split(FuncIDtemp,"&&")
	FuncStatus=0
	for qqqq=0 to ubound(FunctionTemp)
		ATemp=split(trim(FunctionTemp(qqqq)),",")
		'response.write FID&ATemp(0)&","&FuncStatus&"<br>"
		if trim(ATemp(0))=trim(FID) then
			FuncStatus=1
			exit for
			'response.write FID&ATemp(0)&"<br>"
		end if
		
	next
	if FuncStatus=0 then
		Response.Redirect "/traffic/Traffic_Login.asp?Error=礚ㄏノセ╰参ぇ舦"
	end if
end function
'浪琩琌Τ琩高穝糤单舦
public function CheckPermission( FunctionID , ActionID ) 
	'ActionID 琩高:1
	'		  穝糤:2
	'		  э:3
	'		  埃:4
	FunctionTemp=split(FuncIDtemp,"&&")
	FuncStatus=0
	for qqq=0 to ubound(FunctionTemp)
		ATemp=split(trim(FunctionTemp(qqq)),",")
		if trim(ATemp(0))=trim(FunctionID) then
			if ATemp(trim(ActionID))="1" then
				CheckPermission=true
			else
				CheckPermission=false
			end if
		end if
	next
end function
%>