<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	set WShShell = Server.CreateObject("WScript.Shell")
'	WShShell.Run Server.MapPath("\wkhtmltopdf") & "\clear.bat",1,true

'	Set ExecutorClear = Server.CreateObject("ASPExec.Execute")   
'	ExecutorClear.Application = Server.MapPath("\wkhtmltopdf") & "\clear.bat"
'	strResult = ExecutorClear.ExecuteWinApp
'	Set ExecutorClear = Nothing


'	Set Executor = Server.CreateObject("ASPExec.Execute")  
'
'	Executor.Application = Server.MapPath("\wkhtmltopdf") & "\wkhtmltoimage.exe"  '指定要執行的應用程式路徑

	Set fso = CreateObject("Scripting.FileSystemObject")

	ServerIP=request.servervariables("LOCAL_ADDR")
	BillSN=split(Request("Jpg_SendBillSN"),",")

	dim Ms_Conn
	Set Ms_Conn = Server.CreateObject("ADODB.Connection")
	Provider="Provider=SQLOLEDB;Data Source=10.112.1.194;initial Catalog=traffic;User Id=sa;Password=#Mitac123;"
	Ms_Conn.Open Provider

	For i = 0 to Ubound(BillSN)

		strSQL="select BillNo,(Case When memberstation='3N00' then 'C' when memberstation='3O00' then 'D'" & _
		" when memberstation='3P00' then 'E' when memberstation='3R00' then 'F'" & _
		" else 'G' end) mStation," & _
		"(Forfeit1+nvl(Forfeit2,0)) Forfeit," & _
		"(Select JudeDate from PasserJude where BillSN=PasserBase.sn) JudeDate," & _
		"(Select JudeDate+30 from PasserJude where BillSN=PasserBase.sn) DealDate," & _
		"(Select (to_Number(to_Char(JudeDate,'YYYY'))-1911||to_char(JudeDate,'MM')) from PasserJude where BillSN=PasserBase.sn) listDate," & _
		"(Select WordNum from UnitInfo where UnitID=PasserBase.memberstation) WordNum " & _
		" from PasserBase where sn="&BillSN(i)

		set ch_rs=conn.execute(strSQL)
		
		CaseNo=""

		strSQL="select Case_No from trat001 where VL_BIL_No='"&ch_rs("BillNo")&"'"
		set ms_rs=Ms_Conn.execute(strSQL)

		CaseNo=trim(ms_rs("Case_No"))

		ms_rs.close

		If not ch_rs.eof Then
			
			cntfile=0

			ms_sql="select count(1) cnt from exec_books where bureau_no='"&ch_rs("mStation")&"' and book_title='"& ch_rs("mStation") & ch_rs("listDate") &"'"
			set ms_rs=Ms_Conn.execute(ms_sql)

			cntfile=cdbl(ms_rs("cnt"))

			ms_rs.close

			If cntfile = 0 Then
				ms_sql="insert into exec_books(bureau_no,book_title)" & _
				" values('"&ch_rs("mStation")&"','"& ch_rs("mStation") & ch_rs("listDate") &"')"

				Ms_Conn.execute(ms_sql)
			
			End if 

			cntfile=0

			ms_sql="select count(1) cnt from exec_certs where case_no='"&CaseNo&"'"
			set ms_rs=Ms_Conn.execute(ms_sql)

			cntfile=cdbl(ms_rs("cnt"))

			ms_rs.close

			If cntfile = 0 Then

				ms_sql="insert into exec_certs(exec_book_id,case_no,cert_word,exec_date,exec_due_date,exec_amt)" & _
					"select exec_book_id," & _
					"'"& CaseNo &"' case_no,'"& ch_rs("WordNum") &"' cert_word,'"& gInitDT(ch_rs("JudeDate")) &"' exec_date," & _
					"'"& gInitDT(ch_rs("DealDate")) &"' exec_due_date,"& ch_rs("Forfeit") &" exec_amt from exec_books where bureau_no='"&ch_rs("mStation")&"' and book_title='"& ch_rs("mStation") & ch_rs("listDate") &"'"

				Ms_Conn.execute(ms_sql)
			end If 
		
		End if 
		ch_rs.close

		strSQL="select BillSN,BillNo,JudeDate from PasserJude where BillSN="&BillSN(i)
		set rs=conn.execute(strSQL)

		If fso.FolderExists(Server.mappath("\img\"&gInitDT(rs("JudeDate")))) = False Then
			fso.createFolder(Server.mappath("\img\"&gInitDT(rs("JudeDate"))))
		End If 
		
		If Not Fso.FileExists(Server.MapPath("\img\"&gInitDT(rs("JudeDate"))&"\"&CaseNo&".jpg")) then
			
			pathstr=" --zoom 1.0 --crop-w 800 http://"&ServerIP&"/traffic/PasserBase/PasserJudeBatList_Pdf_miaoli.asp?Sys_PasserJude=1&Sys_SendBillSN="&BillSN(i)&"&JpgUnitID="&Session("Unit_ID")&" "&Server.MapPath("\img\"&gInitDT(rs("JudeDate"))&"\"&CaseNo&".jpg")
			
			WShShell.Run Server.MapPath("\wkhtmltopdf") & "\wkhtmltoimage.exe"&pathstr,1,true

			'Executor.Parameters = "--zoom 1.0 --crop-w 800 http://"&ServerIP&"/traffic/PasserBase/PasserJudeBatList_Pdf_miaoli.asp?Sys_PasserJude=1&Sys_SendBillSN="&BillSN(i)&"&JpgUnitID="&Session("Unit_ID")&" "&Server.MapPath("\img\"&gInitDT(rs("JudeDate"))&"\"&CaseNo&".jpg")  '執行應用程式所需的參數                      

			'strResult = Executor.ExecuteDosApp        
		end if
		
		rs.close
	Next

	Set fso = Nothing
    
	Set Executor = Nothing

	Ms_Conn.close
	set Ms_Conn=nothing
	
'	WShShell.Run Server.MapPath("\wkhtmltopdf") & "\clear.bat",1,true
'	Set ExecutorClear = Server.CreateObject("ASPExec.Execute")   
'	ExecutorClear.Application = Server.MapPath("\wkhtmltopdf") & "\clear.bat"
'	strResult = ExecutorClear.ExecuteWinApp
'	Set ExecutorClear = Nothing

	Response.write "<script>"
	Response.Write "alert('轉換完成！！');"
	response.write "self.close();"
	Response.write "</script>"
%>