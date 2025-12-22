<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT> 
dim Data_5xsoft 

Class upload_5xsoft 
	
	dim objForm,objFile,Version 
	'------------------------begin add by mytju.com---------------------------- 
	dim mProgressID,tempFileName 
	Public Property Get ProgressID() 
		ProgressID=mProgressID 
	End Property 

	Public Property Let ProgressID(byVal itemValue) 
		mProgressID=itemValue 
	End Property 
	'-----------------------end add by mytju.com----------------------------- 
	Public function Form(strForm) 
		strForm=lcase(strForm) 
		if not objForm.exists(strForm) then
			Form="" 
		else 
			Form=objForm(strForm) 
		end if 
	end function 

	Public function File(strFile) 
		strFile=lcase(strFile) 
		if not objFile.exists(strFile) then 
			set File=new FileInfo 
		else 
			set File=objFile(strFile) 
		end if 
	end function 


	Public Sub GetUpFile '-----------change by mytju.com------------------------------ 
		dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile 
		dim iFileSize,sFilePath,sFileType,sFormValue,sFileName 
		dim iFindStart,iFindEnd 
		dim iFormStart,iFormEnd,sFormName 
		dim readBlock,readBlockSize,upBytes 
		Version="化境HTTP上傳程序Version 2.0" 
		set objForm=Server.CreateObject("Scripting.Dictionary") 
		set objFile=Server.CreateObject("Scripting.Dictionary") 
		if Request.TotalBytes<1 then Exit Sub 
		set tStream = Server.CreateObject("adodb.stream") 
		set Data_5xsoft = Server.CreateObject("adodb.stream") 
		Data_5xsoft.Type = 1 
		Data_5xsoft.Mode =3 
		Data_5xsoft.Open 
	'------------------------modify begin by mytju.com------------------------------ 
	'---create xml File-------------------- 
		Dim objFSO, objTextFile 
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
		tempFileName=objFSO.GetSpecialFolder(2)&"\upload_"&mProgressID&".xml" 
		Set objTextFile = objFSO.CreateTextFile(tempFileName, True) 
		objTextFile.WriteLine("<?xml version=""1.0"" encoding=""big5""?>") 
		objTextFile.WriteLine("<tree>") 
		objTextFile.WriteLine("<fileUpProgress myID=""1"">") 
		objTextFile.WriteLine("<BytesTransferred>0</BytesTransferred>") 
		objTextFile.WriteLine("<TotalBytes>0</TotalBytes>") 
		objTextFile.WriteLine("<useTime>1</useTime>") 
		objTextFile.WriteLine("</fileUpProgress>") 
		objTextFile.WriteLine("</tree>") 
		objTextFile.Close 
		Set objTextFile = Nothing 
		Set objFSO = Nothing 
	 '---start to upload--------------------- 
		upBytes=0 
		readBlockSize=1024*100 
		startTime=Timer() 
		readBlock=Request.BinaryRead(readBlockSize) 
	 '---create xmlDom------------------------ 
		set objDOM = Server.CreateObject("Microsoft.XMLDOM") 
		objDOM.async = false 
		objDOM.load tempFileName 
		Set objRoot = objDom.documentElement 
		Set objField = objRoot.selectSingleNode("/tree/fileUpProgress[@myID=1]") 
	 '---loop to get data-------------------- 
		while Lenb(readBlock)>0 
			upBytes=upBytes+Lenb(readBlock) 
			'--save value--------- 
			objField.childNodes.item(0).text=Cstr(upBytes) 
			objField.childNodes.item(1).text=Cstr(Request.TotalBytes) 
			objField.childNodes.item(2).text=Cstr(Timer()-startTime) 
			objDom.save tempFileName 
			'--write to Stream----- 
			Data_5xsoft.Write readBlock 
			readBlock=Request.BinaryRead(readBlockSize) 
		wend 
		Set objField = Nothing 
		Set objRoot = Nothing 
		Set objDom = Nothing 
	'------------------------modify end by mytju.com------------------------------ 
		Data_5xsoft.Position=0 
		RequestData =Data_5xsoft.Read 

		iFormStart = 1 
		iFormEnd = LenB(RequestData) 
		vbCrlf = chrB(13) & chrB(10) 
		sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1) 
		iStart = LenB (sStart) 
		iFormStart=iFormStart+iStart+1 
		while (iFormStart + 10) < iFormEnd 
			iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3 
			tStream.Type = 1 
			tStream.Mode =3 
			tStream.Open 
			Data_5xsoft.Position = iFormStart 
			Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart 
			tStream.Position = 0 
			tStream.Type = 2 
			tStream.Charset ="big5" 
			sInfo = tStream.ReadText 
			tStream.Close 
			'取得表單項目名稱
			iFormStart = InStrB(iInfoEnd,RequestData,sStart) 
			iFindStart = InStr(22,sInfo,"name=""",1)+6 
			iFindEnd = InStr(iFindStart,sInfo,"""",1) 
			sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart)) 
			'如果是文件
			if InStr (45,sInfo,"filename=""",1) > 0 then 
				set theFile=new FileInfo 
				'取得文件名 
				iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10 
				iFindEnd = InStr(iFindStart,sInfo,"""",1) 
				sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart) 
				theFile.FileName=getFileName(sFileName) 
				theFile.FilePath=getFilePath(sFileName) 
				'取得文件類型 
				iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14 
				iFindEnd = InStr(iFindStart,sInfo,vbCr) 
				theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart) 
				theFile.FileStart =iInfoEnd 
				theFile.FileSize = iFormStart -iInfoEnd -3 
				theFile.FormName=sFormName 
				if not objFile.Exists(sFormName) then 
					objFile.add sFormName,theFile 
				end if 
			else 
				'如果是表單項目 
				tStream.Type =1 
				tStream.Mode =3 
				tStream.Open 
				Data_5xsoft.Position = iInfoEnd 
				Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3 
				tStream.Position = 0 
				tStream.Type = 2 
				tStream.Charset ="big5" 
				sFormValue = tStream.ReadText 
				tStream.Close 
				if objForm.Exists(sFormName) then 
					objForm(sFormName)=objForm(sFormName)&","&sFormValue		
				else 
					objForm.Add sFormName,sFormValue 
				end if 
			end if 
			iFormStart=iFormStart+iStart+1 
		wend 
		RequestData="" 
		set tStream =nothing 
	End Sub 

	Private Sub Class_Terminate 
		if Request.TotalBytes>0 then 
			objForm.RemoveAll 
			objFile.RemoveAll 
			set objForm=nothing 
			set objFile=nothing 
			Data_5xsoft.Close 
			set Data_5xsoft =nothing 
		end if 
	End Sub 
		 

	Private function GetFilePath(FullPath) 
		If FullPath <> "" Then 
			GetFilePath = left(FullPath,InStrRev(FullPath, "\")) 
		Else 
			GetFilePath = "" 
		End If 
	End function 

	Private function GetFileName(FullPath) 
		If FullPath <> "" Then 
			GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1) 
		Else 
			GetFileName = "" 
		End If 
	End function 
End Class 

Class FileInfo 
	dim FormName,FileName,FilePath,FileSize,FileType,FileStart 

	Private Sub Class_Initialize 
		FileName = "" 
		FilePath = "" 
		FileSize = 0 
		FileStart= 0 
		FormName = "" 
		FileType = "" 
	End Sub 
	
	Public function SaveAs(FullPath) 
		dim dr,ErrorChar,i 
		SaveAs=true 
		if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function 
		set dr=CreateObject("Adodb.Stream") 
		dr.Mode=3 
		dr.Type=1 
		dr.Open 
		Data_5xsoft.position=FileStart 
		Data_5xsoft.copyto dr,FileSize 
		dr.SaveToFile FullPath,2 
		dr.Close 
		set dr=nothing 
		SaveAs=false 
	end function 
End Class 
</SCRIPT>