<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>慢車掃描批次處理系統</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.style2 {font-size: 10px; }
-->
</style>
</HEAD>
<%

Server.ScriptTimeout=6000

'strCity="select value from Apconfigure where id=31"
'set rsCity=conn.execute(strCity)
'sys_City=trim(rsCity("value"))
'rsCity.close

set WShShell = Server.CreateObject("WScript.Shell")

sscannerDir=Session("Credit_ID")

fp=server.mappath(".\PBFtp\PasserFtp")&"\"&scannerDir

mdir=server.mappath(".\PasserImage")&"\"

UpFilePath="./Picture/"

upDir=server.mappath(".\Picture")&"\"

cf_FileName = server.mappath(".\PBFtp\Temp")&"\A"&Session("User_ID")&".bat"

Set cf = fso.CreateTextFile(cf_FileName , true)

set fso=Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(fp) Then fso.CreateFolder (fp)
set fod=fso.GetFolder(fp)
set fic=fod.Files

For Each fil In fic
	if instr(",JPG,JPEG,PDF,",","&UCase(fso.GetExtensionName(fil.Name))&",") >0 then

		fname=UCase(trim(fil.Name))

		strSQL="delete PasserImage where Imagefilename='"&fname&"'"
		conn.execute(strSQL)

		if left(fname,2)="J0" then

			PBillsn=replace(fname,"J0","")

			'1裁決 2催告 3移送

			strSQL="insert into PasserImage(BillSN,PkeySN,ImgKindID,ImgTypeID,ImageFileName,RecordDate,RecordMemberID) "&_
			" values("&PBillsn&","&PBillsn&",1,1,'"&fname&"',sysdate,"&Session("User_ID")&")"
			conn.execute(strSQL)

			patname="move /y "&fp & "\" & fname &" "&mdir & fname

			cf.WriteLine(patname)

		elseif left(fname,2)="S0" then

			PBillsn=replace(fname,"S0","")

			'1裁決 2催告 3移送

			sys_BillSN=""

			strSQL="select BillSN from PasserSendDetail where sn="&PBillsn
			set rspay=conn.execute(strSQL)
			If not rspay.eof Then sys_BillSN=cdbl(rspay("BillSN"))
			rspay.close

			If sys_BillSN <>"" Then

				strSQL="insert into PasserImage(BillSN,PkeySN,ImgKindID,ImgTypeID,ImageFileName,RecordDate,RecordMemberID) "&_
				" values("&sys_BillSN&","&PBillsn&",3,1,'"&fname&"',sysdate,"&Session("User_ID")&")"

				conn.execute(strSQL)

				patname="move /y "&fp & "\" & fname &" "&mdir & fname

				cf.WriteLine(patname)
			
			End if 

		elseif left(fname,2)="D0" then

			PBillsn=replace(fname,"D0","")

			'1裁決 2催告 3移送

			strSQL="select Max(SN) as cnt from PassersEndArrived"
			set rscnt=conn.execute(strSQL)
			PasserSN=1
			if Not isnull(rscnt("cnt")) then
				PasserSN=cdbl(rscnt("cnt"))+1
			end if
			rscnt.close

			strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,SendMailStation,ArriveType,ReturnResonID,Note) values("&PasserSN&"(select billsn from passerbase where billno='"&PBillsn&"' and recordstateid=0),null,null,"&Session("User_ID")&",null,null,null,null)"
			
			conn.execute(strSQL)

			strSQL="Update PassersEndArrived set Imagedirectoryname='"&UpFilePath&"',Imagefilename='"&fname&"' where SN="&PasserSN

			conn.execute(strSQL)

			patname="move /y "& fp & "\" & fname &" "& upDir & fname

			cf.WriteLine(patname)

		end If 

	end if
Next

cf.WriteLine("exit")
cf.close
set cf=nothing

WShShell.Run server.mappath(".\PBFtp\Temp")&"\A"&Session("User_ID")&".bat",1,true

set fso=nothing

Response.write "<script>"
Response.Write "alert('儲存完成！');"
Response.write "</script>"

%>
<BODY>
<form name=myForm method="post">

</form>
</BODY>
</HTML>
