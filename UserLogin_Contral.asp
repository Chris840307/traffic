<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>使用者登入</title>
</head>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<%
Session.Contents.Remove("FuncID")
Session.Contents.Remove("User_ID")
Session.Contents.Remove("Unit_ID")
Session.Contents.Remove("Ch_Name")
Session.Contents.Remove("Credit_ID")
Session.Contents.Remove("Group_ID")
Session.Contents.Remove("DoubleCheck")
Session.Contents.Remove("UnitLevelID")
Session.Contents.Remove("ManagerPower")
Session.Contents.Remove("DCIwindowName")
Session.Contents.Remove("BillIgnore_Fix")
Session.Contents.Remove("BillIgnore_Image")
Session.Contents.Remove("SpecUser")
Session.Contents.Remove("UpdateBillUser")
Session.Contents.Remove("RuleVer")

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	user_id=Request("MemberID")
	user_id=Replace(user_id,"'", "")
	user_pwd=Request("MemberPW")
	If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" Then
		user_pwd=decrypt(Replace(user_pwd,"'", ""))
	Else
		user_pwd=Replace(user_pwd,"'", "")
	End If   	
  	if trim(user_id)<>"" and trim(user_pwd)<>"" then
		strSql="select * from MemberData where AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null and CreditID='"&trim(user_id)&"' and password='"&user_pwd&"'"
	elseif trim(request("PKICarchk"))<>"" then
		strSQL="select * from MemberData where AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null and PKI='"&trim(request("PKICarchk"))&"'"
	end if
	set rs1=conn.execute(strSql)
	if rs1.eof then
		if trim(user_id)<>"" and trim(user_pwd)<>"" Then
			str1="select * from memberdata where AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null " &_
				" and CreditID='"&trim(user_id)&"'"
			Set rs1=conn.execute(str1)
			If Not rs1.eof Then
				'寫入LOG
				strI="insert into Log(Sn,TypeID,ActionMemberID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),357,"&Trim(rs1("MemberID"))&",'密碼錯誤','"&userip&"',sysdate,'"&user_id&" "&"登入密碼錯誤')"
				Conn.execute strI
				
		'If sys_City="台南市" then
				ECount=0
				ExitFlag=0
				str2="select TypeID From (" &_
					"select TypeID from Log where ActionMemberID="&Trim(rs1("MemberID"))&" and ActionChName not in ('帳號錯誤','權限異常') and ActionDate>(sysdate-60) order by ActionDate desc" &_
					") where Rownum<=3"
				Set rs2=conn.execute(str2)
				While Not rs2.Eof
					If ExitFlag=0 then
						If Trim(rs2("TypeID"))<>"357" then
							ExitFlag=1
						Else	
							ECount=ECount+1
						End If 
					End if
				rs2.MoveNext
				Wend
				rs2.close
				Set rs2=Nothing 
				
				If ECount=1 Then
					WECount="１"
				ElseIf ECount=2 Then
					WECount="２"
				ElseIf ECount=3 Then
					WECount="３"
				End If 

				If ECount=3 Then
					'鎖定帳號
					str3="Update Memberdata set AccountStateID=-1,DelMemberID=99999,LeaveJOBDate=sysdate where AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null and CreditID='"&trim(user_id)&"'"
					conn.execute str3
					'寫log
					strI="insert into Log(Sn,TypeID,ActionMemberID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),358,"&Trim(rs1("MemberID"))&",'"&Trim(rs1("ChName"))&"','"&userip&"',sysdate,'"&user_id&" "&"密碼輸入錯誤3次，帳號封鎖')"
					Conn.execute strI

					Response.Redirect "Traffic_Login.asp?UpdM=0&ErrorS=密碼輸入錯誤"&WECount&"次，需由單位系統管理人員再度啟用帳號"
				Else
					Response.Redirect "Traffic_Login.asp?UpdM=0&ErrorS=密碼錯誤請查明後再輸入，已輸入錯誤 "&WECount&" 次，連續輸入錯誤 ３ 次，系統會將帳號封鎖"
				End If 				
'		Else
'				Response.Redirect "Traffic_Login.asp?ErrorS=密碼錯誤請查明後再輸入"
'		End If 
			End If 
				'寫入LOG
				strI="insert into Log(Sn,TypeID,ActionMemberID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),357,Null,'帳號錯誤','"&userip&"',sysdate,'"&user_id&" "&"登入帳號錯誤')"
				Conn.execute strI
				Response.Redirect "Traffic_Login.asp?UpdM=0&ErrorS=帳號錯誤請查明後再輸入，如帳號被鎖定，需由分局或交通隊系統管理人員再度啟用帳號"
			rs1.close 
			Set rs1=Nothing 
			
		elseif trim(request("PKICarchk"))<>"" then
			Response.Redirect "Traffic_Login.asp?UpdM=0&ErrorS=無此自然人憑證，請輸入帳號密碼再以自然人憑證登入"
		end if
	elseif trim(rs1("GrouproleID"))="0" or isnull(rs1("GrouproleID")) or trim(rs1("GrouproleID"))="" Then
		'寫入LOG
		strI="insert into Log(Sn,TypeID,ActionMemberID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),357,Null,'權限異常','"&userip&"',sysdate,'"&user_id&" 登入權限異常')"
		Conn.execute strI
		Response.Redirect "Traffic_Login.asp?UpdM=0&ErrorS=您目前沒有權限，請通知帳號管理員為您加入權限" 
	Else
	'=檢查三個月未登入使用者(每週五14~15時執行)====================
	'If sys_City<>"台中縣" And sys_City<>"高雄縣" and sys_City<>"台南縣" Then
	If sys_City="澎湖縣" or sys_City="苗栗縣" Then
	If Weekday(date)=6 And Hour(now)>11 And Hour(now)<13 then
		chkLogDate=Year(DateAdd("m",-3,date))&"/"&month(DateAdd("m",-3,date))&"/"&day(DateAdd("m",-3,date))
		strNoLogUser="select * from Memberdata where accountstateid=0 and recordstateid=0 " &_
			" and password is not null and GroupRoleID<>0"
		Set rsNLU=conn.execute(strNoLogUser)
		If Not rsNLU.eof Then rsNLU.MoveFirst
		While Not rsNLU.Eof
			strNoLog="select count(*) as cnt from Log where ActionDate > to_date('"&chkLogDate&"','yyyy/mm/dd') " &_
				" and ActionMemberID="&Trim(rsNLU("MemberID"))&" and TypeID=350"
			Set rsNoLog=conn.execute(strNoLog)
			If Not rsNoLog.eof Then
				If CDbl(rsNoLog("cnt")) = 0 Then
					strMemUpd="Update Memberdata set PassWord=null,GroupRoleID=0 " &_
						" where memberid="&Trim(rsNLU("MemberID"))
					
					conn.execute strMemUpd

					'寫log
					strI="insert into Log(Sn,TypeID,ActionMemberID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),358,"&Trim(rsNLU("MemberID"))&",'"&Trim(rsNLU("ChName"))&"','"&userip&"',sysdate,'"&Trim(rsNLU("CreditID"))&" "&"超過三個月未登入，帳號封鎖')"
					Conn.execute strI
					'response.write strI & "<br>"
				End if
			End if
			rsNoLog.close
			Set rsNoLog=Nothing
			'response.write strNoLog & "<br>"
		rsNLU.MoveNext
		Wend
		rsNLU.close
		Set rsNLU=Nothing 
	End If
	End if
	'=====================================
		'將權限存入session
		if trim(request("PKICarchk"))<>"" and trim(user_id)<>"" and trim(user_pwd)<>"" then
			strSQL="Update MEMBERDATA set PKI='"&trim(request("PKICarchk"))&"' where CreditID='"&trim(user_id)&"' and AccountStateID=0 and RecordstateID=0 and LeaveJOBDate is null and password='"&user_pwd&"'"
			conn.execute(strSQL)
		end if
		strFunc1="select * from functionDataDetail where GroupID='"&trim(rs1("GrouproleID"))&"'"
		set rsFunc1=conn.execute(strFunc1)
		If Not rsFunc1.eof Then rsFunc1.MoveFirst
		While Not rsFunc1.Eof
			if FuncTemp="" then
				FuncTemp=trim(rsFunc1("systemID"))&","&trim(rsFunc1("SelectFlag"))&","&trim(rsFunc1("InsertFlag"))&","&trim(rsFunc1("UpdateFlag"))&","&trim(rsFunc1("DeleteFlag"))
			else
				FuncTemp=FuncTemp&"&&"&trim(rsFunc1("systemID"))&","&trim(rsFunc1("SelectFlag"))&","&trim(rsFunc1("InsertFlag"))&","&trim(rsFunc1("UpdateFlag"))&","&trim(rsFunc1("DeleteFlag"))
			end if
		rsFunc1.MoveNext
		Wend
		rsFunc1.close
		set rsFunc1=nothing
		'response.write FuncTemp
		Session("FuncID")=FuncTemp
		Session("Unit_ID")=rs1("UnitID")
		Session("User_ID")=rs1("MemberID")
		Session("Ch_Name")=rs1("ChName")
		Session("Credit_ID")=rs1("CreditID")
		Session("Group_ID")=rs1("GroupRoleID")
		Session("ManagerPower")=rs1("ManagerPower")
		'單位等級
		strUnit="select UnitLevelID,DCIwindowName from UnitInfo where UnitID='"&trim(rs1("UnitID"))&"'"
		set rsUnit=conn.execute(strUnit)
		if not rsUnit.eof then
			UnitLevel=trim(rsUnit("UnitLevelID"))
			DCIwindow=trim(rsUnit("DCIwindowName"))
		end if
		rsUnit.close
		set rsUnit=nothing
		Session("UnitLevelID")=UnitLevel
		Session("DCIwindowName")=DCIwindow

		'一打一驗判斷
		DoubleChk=0
		strDoubleChk="select ID,Value from Apconfigure where ID in (3,38)"
		set rsDbChk=conn.execute(strDoubleChk)
		if not rsDbChk.eof Then
			If trim(rsDbChk("ID"))="38" Then 
				DoubleChk=trim(rsDbChk("Value"))
			Else
				Session("RuleVer")=trim(rsDbChk("Value"))
			End If 
		end if
		rsDbChk.close
		set rsDbChk=nothing
		Session("DoubleCheck")=DoubleChk
		
		'讀取特定使用者(業管車輛)
		Session("SpecUser")="0"
		strSpecUser=""
		dim fso 
		set fso=Server.CreateObject("Scripting.FileSystemObject")
		SpecuserFile = Server.MapPath("/traffic/Common/SpecUser.ini")	' 找出記數檔案在硬碟中的實際位置
		Set Out= fso.OpenTextFile(SpecuserFile, 1, FALSE)	' 開啟唯讀檔案
		if not Out.atEndOfStream then
			strSpecUser = Out.ReadLine	' 從檔案讀出記數資料
		end if

		Out.Close	' 關閉檔案
		set fso=nothing

		if strSpecUser<>"" then
			SUserArray=split(strSpecUser,",")
			for SU=0 to ubound(SUserArray)
				if SUserArray(SU)=Session("Credit_ID") then
					Session("SpecUser")="1"
					exit for
				end if
			next
	
		end if

		'讀取特定使用者(整批修改)
		Session("UpdateBillUser")="0"
		strUpdateBillUser=""
		dim fso1 
		set fso1=Server.CreateObject("Scripting.FileSystemObject")
		UpdateBillFile = Server.MapPath("/traffic/Common/UpdateBillUser.ini")	' 找出記數檔案在硬碟中的實際位置
		Set Out1= fso1.OpenTextFile(UpdateBillFile, 1, FALSE)	' 開啟唯讀檔案
		if not Out1.atEndOfStream then
			strUpdateBillUser = Out1.ReadLine	' 從檔案讀出記數資料
		end if

		Out1.Close	' 關閉檔案
		set fso1=nothing

		if strUpdateBillUser<>"" then
			UpdBUserArray=split(strUpdateBillUser,",")
			for SU=0 to ubound(UpdBUserArray)
				if UpdBUserArray(SU)=Session("Credit_ID") then
					Session("UpdateBillUser")="1"
					exit for
				end if
			next
	
		end if
		'response.write trim(Session("SpecUser"))
		'response.write Session("FuncID")&"<br>"&Session("Unit_ID")&"<br>"&Session("User_ID")&"<br>"&Session("Ch_Name")&"<br>"&Session("Credit_ID")&"<br>"&Session("Group_ID")&"<br>"&Session("DoubleCheck")
		'寫入LOG
		ServerIp999 = Request.ServerVariables("Local_ADDR") 

		ConnExecute Session("Credit_ID")&","&Session("Ch_Name")&" "&"登入" & ServerIp999,350 
		if UseIpSelectUrlLocationFlag=1 then
		'寫入登入的主機的ip

			
			Serverlogsql="Insert Into Log(SN,TypeID,ActionMemBerID,ActionChName,ActionIP,ActionDate,ActionContent) values((select nvl(max(Sn),0)+1 from Log),999,"&Session("User_ID")&",'"&Session("Ch_Name") &"','"&ServerIp999&"',"&funGetDate(now,1)&",'"&Session("Credit_ID")&" 登入主機 "&ServerIp999&"')"
			conn.execute(Serverlogsql)
		end If
		
		'------------------------------------------
		ErrorNo=0
		ErrorString=""
		Modifydate=""
		PassWordTemp=""
		showUpdateMemberdateFlag=0
		strMem="select * from memberdata where memberid=" & Trim(session("User_ID")) & " and recordstateid=0 and accountstateid=0"
		Set rsMem=conn.execute(strMem)
		If Not rsMem.eof Then
			If Trim(rsMem("ModifyTime") & "")<>"" Then
				Modifydate=Trim(rsMem("ModifyTime"))
			Else
				Modifydate=Trim(rsMem("RecordDate"))
			End If 
			If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" Then
				PassWordTemp=encrypt(Trim(rsMem("PassWord")))
			Else
				PassWordTemp=Trim(rsMem("PassWord"))
			End If 			
		End If 
		rsMem.close
		Set rsMem=Nothing 
		If DateDiff("d",Modifydate,now)>90 Then
			ErrorNo=ErrorNo+1
			ErrorString=ErrorString & "\n" & ErrorNo & ":您的密碼已超過三個月( " & DateDiff("d",Modifydate,now) & " 天)未更換。"
			showUpdateMemberdateFlag=1
		End If 
		if Len(PassWordTemp)<8 Then
			ErrorNo=ErrorNo+1
			ErrorString=ErrorString & "\n" & ErrorNo & ":密碼長度至少為8碼，且需包含英文、數字、特殊符號及大小寫混和。"
			showUpdateMemberdateFlag=1
		End If 
		If showUpdateMemberdateFlag=0 then
			chkUp=0
			chkDown=0
			chkInt=0
			chkMark=0
			for i=1 to Len(PassWordTemp)
				if Asc(Mid(Trim(PassWordTemp), i, 1))>=65 and Asc(Mid(Trim(PassWordTemp), i, 1))<=90 then
					chkUp=1
				end if
				if Asc(Mid(Trim(PassWordTemp), i, 1))>=97 and Asc(Mid(Trim(PassWordTemp), i, 1))<=122 then
					chkDown=1
				end if 
				if Asc(Mid(Trim(PassWordTemp), i, 1))>=48 and Asc(Mid(Trim(PassWordTemp), i, 1))<=57 then
					chkInt=1
				end if 
				if (Asc(Mid(Trim(PassWordTemp), i, 1))>=33 and Asc(Mid(Trim(PassWordTemp), i, 1))<=47) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=58 and Asc(Mid(Trim(PassWordTemp), i, 1))<=64) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=91 and Asc(Mid(Trim(PassWordTemp), i, 1))<=96) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=123 and Asc(Mid(Trim(PassWordTemp), i, 1))<=126) then
					chkMark=1
				end if
			next
			if chkUp=0 or chkDown=0 or chkInt=0 or chkMark=0 Then
				ErrorNo=ErrorNo+1
				ErrorString=ErrorString & "\n" & ErrorNo & ":密碼長度至少為8碼，且需包含英文、數字、特殊符號及大小寫混和。"
				'response.write "alert('請儘速修改您的密碼!!');"
				showUpdateMemberdateFlag=1
			end if
		End If 

		If showUpdateMemberdateFlag=1 Then
			Response.Redirect "Traffic_Login.asp?UpdM=1&ErrorS="&ErrorString&"\n請修改密碼後再次登入!!"
		else
			If sys_City="台南市" then
				Response.Redirect "Traffic_Web_Main_TN.asp"
			Else
				Response.Redirect "Traffic_Web_Main.asp"
			End if
		End If 

		
	end if
	rs1.close
	set rs1=nothing
conn.close
set conn=nothing
%>
<body>
</body>
</html>
