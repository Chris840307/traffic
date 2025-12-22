<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title></title>
</head>
<%
Session.TimeOut=30
Session.Contents.Remove("User_Unit_ID")
Session.Contents.Remove("User_ID")
Session.Contents.Remove("User_Logined")
Session.Contents.Remove("FuncFAX")
Session.Contents.Remove("FuncContact")
Session("User_Logined")="false"
dim conn
Set conn = Server.CreateObject("ADODB.Connection")  
conn.Open "dsn=TFD;uid=TFD;pwd=tfduser;"
	user_id=Request.Form("MemberID")
	user_id=Replace(user_id,"'", "")	
	user_pwd=Request.Form("MemberPW")
  	user_pwd=Replace(user_pwd,"'", "")
  	
	strSql="select * from MemberData where MemID='"&trim(user_id)&"'"
	set rs1=conn.execute(strSql)
	if rs1.eof then
		Response.Redirect "TFD_Login.asp?Error=無此帳號" 
	elseif trim(rs1("GroupID"))="0" then
		Response.Redirect "TFD_Login.asp?Error=您目前沒有權限，請通知帳號管理員為您加入權限" 
	else
		if trim(user_pwd)=trim(rs1("Pwd")) then
			response.cookies("MemID")=trim(user_id)
			response.cookies("GroupID")=trim(rs1("GroupID"))
			response.cookies("UnitID")=trim(rs1("UnitID"))
		'將權限存入cookie
		strPer="select * from PermissionData where GroupID="&trim(rs1("GroupID"))
		set rsPer=conn.execute(strPer)	
		CountNo=rsPer.fields.count
		i=0
		FuncTemp=""
		for J=2 to CountNo-1
			if trim(rsPer(J))="1" then
				strFunc1="select * from FuncContrast a,PermissionDataDetail b where GroupID="&trim(rs1("GroupID"))&" and a.IsWeb='Y' and a.FuncSN=b.FunctionID and a.FuncID='"&trim(rsPer(J).name)&"'"
				set rsFunc1=conn.execute(strFunc1)
				If Not rsFunc1.eof Then rsFunc1.MoveFirst
				While Not rsFunc1.Eof
					if FuncTemp="" then
						FuncTemp=rsFunc1("FunctionID")&","&rsFunc1("Sel")&","&rsFunc1("Ins")&","&rsFunc1("Upd")&","&rsFunc1("Del")
					else
						FuncTemp=FuncTemp&"&&"&rsFunc1("FunctionID")&","&rsFunc1("Sel")&","&rsFunc1("Ins")&","&rsFunc1("Upd")&","&rsFunc1("Del")
					end if
				rsFunc1.MoveNext
				Wend
				rsFunc1.close
				set rsFunc1=nothing
			end if
		next
		response.write FuncTemp
		rsPer.close
		set rsPer=nothing
			response.cookies("FuncID")=trim(FuncTemp)

			Session("User_Unit_ID")=rs1("UnitID")
			Session("User_ID")=rs1("MemID")
			group_id=rs1("GroupID")
			Session("User_Logined")="true"
			SQL="select * from PermissionData Where GroupID=" & group_id
			set rs=conn.execute(SQL)
			If Not rs.Bof Then rs.MoveFirst
			If Not rs.Eof Then
				Session("FuncFAX")=rs("FuncFax")
				Session("FuncContact")=rs("FuncContact")
			End If 
			strLog="insert into Log(MemID,LogTime,WKS,ActionID,Result,Note,UnitID)"
			strLog=strLog&" values('"&rs1("MemID")&"',getdate(),'WEB系統登入',1318,'1','WEB系統登入','"&trim(rs1("UnitID"))&"')"
			conn.execute strLog
			Response.Redirect "TFD_Main.asp" 
		elseif isnull(rs1("Pwd")) and trim(user_pwd)="" then
			response.cookies("MemID")=trim(user_id)
			response.cookies("GroupID")=trim(rs1("GroupID"))
			response.cookies("UnitID")=trim(rs1("UnitID"))
		'將權限存入cookie
		strPer="select * from PermissionData where GroupID="&trim(rs1("GroupID"))
		set rsPer=conn.execute(strPer)	
		CountNo=rsPer.fields.count
		i=0
		FuncTemp=""
		for J=2 to CountNo-1
			if trim(rsPer(J))="1" then
				strFunc1="select * from FuncContrast a,PermissionDataDetail b where GroupID="&trim(rs1("GroupID"))&" and a.IsWeb='Y' and a.FuncSN=b.FunctionID and a.FuncID='"&trim(rsPer(J).name)&"'"
				set rsFunc1=conn.execute(strFunc1)
				If Not rsFunc1.eof Then rsFunc1.MoveFirst
				While Not rsFunc1.Eof
					if FuncTemp="" then
						FuncTemp=rsFunc1("FunctionID")&","&rsFunc1("Sel")&","&rsFunc1("Ins")&","&rsFunc1("Upd")&","&rsFunc1("Del")
					else
						FuncTemp=FuncTemp&"&&"&rsFunc1("FunctionID")&","&rsFunc1("Sel")&","&rsFunc1("Ins")&","&rsFunc1("Upd")&","&rsFunc1("Del")
					end if
				rsFunc1.MoveNext
				Wend
				rsFunc1.close
				set rsFunc1=nothing
			end if
		next
		response.write FuncTemp
		rsPer.close
		set rsPer=nothing
			response.cookies("FuncID")=trim(FuncTemp)

			Session("User_Unit_ID")=rs1("UnitID")
			Session("User_ID")=rs1("MemID")
			group_id=rs1("GroupID")
			Session("User_Logined")="true"
			SQL="select * from PermissionData Where GroupID=" & group_id
			set rs=conn.execute(SQL)
			If Not rs.Bof Then rs.MoveFirst
			If Not rs.Eof Then
				Session("FuncFAX")=rs("FuncFax")
				Session("FuncContact")=rs("FuncContact")
			End If 
			Response.Redirect "TFD_Main.asp" 
		else
			Response.Redirect "TFD_Login.asp?Error=密碼錯誤Password error" 
		end if
	end if
	
conn.close
set conn=nothing
%>
<body>
</body>
</html>
